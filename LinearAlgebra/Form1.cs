using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LinearAlgebra
{
    public partial class Form1 : Form
    {
        private Random rnd = new Random();

        private bool isBusy = false;   // true, пока идёт загрузка или расчёт

        // последняя посчитанная система и решения (для экспорта)
        private double[,] lastA = null;
        private double[] lastB = null;          // B в исходном виде (AX + B = 0)
        private double[] lastGauss = null;
        private double[] lastGaussJordan = null;
        private double[] lastCramer = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InputAdataGridView.AllowUserToAddRows = false;
            InputBdataGridView.AllowUserToAddRows = false;
            GaussdataGridView.AllowUserToAddRows = false;
            GaussJordandataGridView.AllowUserToAddRows = false;
            CramerdataGridView.AllowUserToAddRows = false;
            CreateTables();
        }

        // при изменении N пересоздаём таблицы
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            CreateTables();
        }

        private void CreateTables()
        {
            int n = (int)numericUpDown1.Value;

            if (n < 2)      // защита: размер меньше 2 не создаём
                return;

            InputAdataGridView.Columns.Clear();
            InputBdataGridView.Columns.Clear();

            // таблица A: n столбцов и n строк
            for (int j = 0; j < n; j++)
            {
                InputAdataGridView.Columns.Add("col" + j, "x" + (j + 1));
                InputAdataGridView.Columns[j].Width = 50;
            }
            InputAdataGridView.Rows.Add(n);

            // таблица B: 1 столбец и n строк
            InputBdataGridView.Columns.Add("colB", "B");
            InputBdataGridView.Columns[0].Width = 50;
            InputBdataGridView.Rows.Add(n);

            // нумерация строк и нули по умолчанию
            for (int i = 0; i < n; i++)
            {
                InputAdataGridView.Rows[i].HeaderCell.Value = (i + 1).ToString();
                InputBdataGridView.Rows[i].HeaderCell.Value = (i + 1).ToString();

                for (int j = 0; j < n; j++)
                    InputAdataGridView.Rows[i].Cells[j].Value = 0;

                InputBdataGridView.Rows[i].Cells[0].Value = 0;
            }
        }

        // Пытается прочитать из ячейки целое число
        // Возвращает false, если ячейка пустая или там не целое число
        private bool TryReadInt(DataGridView grid, int row, int col, out double value)
        {
            value = 0;

            object cell = grid.Rows[row].Cells[col].Value;
            if (cell == null)
                return false;

            int result;
            if (!int.TryParse(cell.ToString().Trim(), out result))
                return false;

            value = result;
            return true;
        }

        // РАСЧЁТ

        private async void рассчитатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isBusy) return;

            // Выбран ли хотя бы один метод?
            if (!GausscheckBox.Checked && !GaussJordancheckBox.Checked && !CramercheckBox.Checked)
            {
                MessageBox.Show("Выберите хотя бы один метод решения!",
                                "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int n = (int)numericUpDown1.Value;
            double[,] A = new double[n, n];
            double[] B = new double[n];

            // Во всех ячейках целые числа?
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    double v;
                    if (!TryReadInt(InputAdataGridView, i, j, out v))
                    {
                        MessageBox.Show("В матрице A (строка " + (i + 1) + ", столбец " + (j + 1) +
                                        ") должно быть целое число!",
                                        "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        InputAdataGridView.CurrentCell = InputAdataGridView.Rows[i].Cells[j];
                        return;
                    }
                    A[i, j] = v;
                }
            }

            // читаем B и меняем знак, т.к. AX + B = 0  =>  AX = -B
            for (int i = 0; i < n; i++)
            {
                double v;
                if (!TryReadInt(InputBdataGridView, i, 0, out v))
                {
                    MessageBox.Show("В векторе B (строка " + (i + 1) + ") должно быть целое число!",
                                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    InputBdataGridView.CurrentCell = InputBdataGridView.Rows[i].Cells[0];
                    return;
                }
                B[i] = -v;
            }

            // Есть ли единственное решение (определитель не равен 0)
            if (Math.Abs(Determinant(A)) < 1e-12)
            {
                MessageBox.Show("Определитель матрицы A равен 0.\n" +
                                "Система не имеет решений или имеет бесконечно много решений.",
                                "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Запоминаем систему для экспорта (B возвращаем в исходный знак)
            lastA = (double[,])A.Clone();
            lastB = new double[n];
            for (int i = 0; i < n; i++)
                lastB[i] = -B[i] + 0.0;
            lastGauss = null;
            lastGaussJordan = null;
            lastCramer = null;

            isBusy = true;   // блокируем остальные кнопки на время расчёта

            // Метод Гаусса
            if (GausscheckBox.Checked)
            {
                Gausslabel.Text = "Считаю...";

                Stopwatch sw = new Stopwatch();
                sw.Start();

                double[] x = await Task.Run(() => Gauss(A, B));

                sw.Stop();

                if (x == null)
                {
                    Gausslabel.Text = "Система не имеет единственного решения";
                    GaussdataGridView.Columns.Clear();
                }
                else
                {
                    ShowResult(GaussdataGridView, x);
                    lastGauss = x;
                    Gausslabel.Text = "Время: " + sw.Elapsed.TotalMilliseconds.ToString("F3") + " мс";
                }
            }

            // Метод Гаусса-Жордана
            if (GaussJordancheckBox.Checked)
            {
                GaussJordanlabel.Text = "Считаю...";

                Stopwatch sw = new Stopwatch();
                sw.Start();

                double[] x = await Task.Run(() => GaussJordan(A, B));

                sw.Stop();

                if (x == null)
                {
                    GaussJordanlabel.Text = "Система не имеет единственного решения";
                    GaussJordandataGridView.Columns.Clear();
                }
                else
                {
                    ShowResult(GaussJordandataGridView, x);
                    lastGaussJordan = x;
                    GaussJordanlabel.Text = "Время: " + sw.Elapsed.TotalMilliseconds.ToString("F3") + " мс";
                }
            }

            // Метод Крамера
            if (CramercheckBox.Checked)
            {
                Cramerlabel.Text = "Считаю...";

                Stopwatch sw = new Stopwatch();
                sw.Start();

                double[] x = await Task.Run(() => Cramer(A, B));

                sw.Stop();

                if (x == null)
                {
                    Cramerlabel.Text = "Система не имеет единственного решения";
                    CramerdataGridView.Columns.Clear();
                }
                else
                {
                    ShowResult(CramerdataGridView, x);
                    lastCramer = x;
                    Cramerlabel.Text = "Время: " + sw.Elapsed.TotalMilliseconds.ToString("F3") + " мс";
                }
            }

            isBusy = false;
        }

        // Метод Гаусса. Возвращает вектор x или null, если решения нет
        private double[] Gauss(double[,] A, double[] B)
        {
            // количество переменных
            int n = B.Length;

            // работаем с копиями, чтобы не портить исходные данные
            double[,] a = (double[,])A.Clone();
            double[] b = (double[])B.Clone();

            // Приводим матрицу к верхнетреугольному виду
            // на шаге k обнуляем все элементы столбца k, расположенные ниже диагонали
            for (int k = 0; k < n; k++)
            {
                // ищем строку с наибольшим по модулю элементом в столбце k
                int maxRow = k;
                for (int i = k + 1; i < n; i++)
                {
                    if (Math.Abs(a[i, k]) > Math.Abs(a[maxRow, k]))
                        maxRow = i;
                }

                // если весь столбец почти нули, решения нет
                if (Math.Abs(a[maxRow, k]) < 1e-12)
                    return null;

                // ставим найденную строку на место строки k
                if (maxRow != k) 
                {
                    for (int j = 0; j < n; j++) // меняем местами строки матрицы
                    {
                        double temp = a[k, j];
                        a[k, j] = a[maxRow, j];
                        a[maxRow, j] = temp;
                    }
                    double tempB = b[k];
                    b[k] = b[maxRow];
                    b[maxRow] = tempB;
                }

                // обнуляем элементы под главным
                for (int i = k + 1; i < n; i++)
                {
                    double m = a[i, k] / a[k, k]; // множитель: во сколько раз строка i "больше" строки k в столбце k
                    for (int j = k; j < n; j++) // вычитаем из строки i строку k, умноженную на m
                                                // (j начинаем с k, т.к. левее k элементы уже равны 0)
                        a[i, j] = a[i, j] - m * a[k, j];
                    b[i] = b[i] - m * b[k];
                }
            }

            // Находим значения переменных снизу вверх 
            double[] x = new double[n];
            for (int i = n - 1; i >= 0; i--)
            {
                double sum = 0; // сумма уже найденных переменных, умноженных на свои коэффициенты
                for (int j = i + 1; j < n; j++)
                    sum = sum + a[i, j] * x[j];

                x[i] = (b[i] - sum) / a[i, i];
            }

            return x;
        }

        // Метод Гаусса-Жордана. Возвращает вектор x или null, если решения нет
        // В отличие от метода Гаусса, приводит матрицу не к треугольному, а сразу к единичному виду,
        // после преобразований решение лежит прямо в векторе b.
        private double[] GaussJordan(double[,] A, double[] B)
        {
            int n = B.Length;

            // работаем с копиями
            double[,] a = (double[,])A.Clone();
            double[] b = (double[])B.Clone();

            // на каждом шаге k "обрабатываем" столбец k:
            // ставим на диагональ единицу, а во всех остальных строках обнуляем элемент этого столбца
            for (int k = 0; k < n; k++)
            {
                // ищем строку с наибольшим по модулю элементом в столбце k
                int maxRow = k;
                for (int i = k + 1; i < n; i++)
                {
                    if (Math.Abs(a[i, k]) > Math.Abs(a[maxRow, k]))
                        maxRow = i;
                }

                if (Math.Abs(a[maxRow, k]) < 1e-12)
                    return null;

                // ставим найденную строку на место строки k
                if (maxRow != k)
                {
                    // меняем местами строки матрицы a
                    for (int j = 0; j < n; j++)
                    {
                        double temp = a[k, j];
                        a[k, j] = a[maxRow, j];
                        a[maxRow, j] = temp;
                    }
                    double tempB = b[k];
                    b[k] = b[maxRow];
                    b[maxRow] = tempB;
                }

                // нормировка: делим всю строку k (и свободный член) на главный элемент,
                // чтобы на диагонали a[k, k] получилась единица
                double main = a[k, k];
                for (int j = 0; j < n; j++)
                    a[k, j] = a[k, j] / main;
                b[k] = b[k] / main;

                // обнуляем столбец k во ВСЕХ остальных строках - и ниже, и выше диагонали
                for (int i = 0; i < n; i++)
                {
                    if (i != k) // строку k не трогаем, она уже нормирована
                    {
                        double m = a[i, k]; // множитель: текущее значение элемента в столбце k строки i
                        for (int j = 0; j < n; j++)
                            // вычитаем из строки i строку k, умноженную на m,
                            // после чего a[i, k] становится равным 0
                            a[i, j] = a[i, j] - m * a[k, j];
                        b[i] = b[i] - m * b[k];
                    }
                }
            }

            // матрица стала единичной, поэтому решение лежит прямо в b
            return b;
        }

        // Метод Крамера. Возвращает вектор x или null, если определитель равен 0
        private double[] Cramer(double[,] A, double[] B)
        {
            int n = B.Length;

            double det = Determinant(A);

            // если определитель почти 0, единственного решения нет
            if (Math.Abs(det) < 1e-12)
                return null;
            
            // Найденные переменные
            double[] x = new double[n];

            // для каждой неизвестной x[j] считаем свой определитель
            for (int j = 0; j < n; j++)
            {
                // делаем копию матрицы A, чтобы не испортить оригинал
                // и заменяем в копии столбец j на столбец свободных членов B
                double[,] temp = (double[,])A.Clone();
                for (int i = 0; i < n; i++)
                    temp[i, j] = B[i];

                // формула Крамера: определитель изменённой матрицы делим на главный определитель
                x[j] = Determinant(temp) / det;
            }

            return x;
        }

        // Определитель квадратной матрицы M.
        // Матрица приводится к верхнетреугольному виду (как в методе Гаусса),
        // а определитель равен произведению элементов диагонали
        // с поправкой на знак за каждую перестановку строк.
        private double Determinant(double[,] M)
        {
            int n = M.GetLength(0);
            // работаем с копией, чтобы не испортить исходную матрицу
            double[,] a = (double[,])M.Clone();
            double det = 1;

            for (int k = 0; k < n; k++)
            {
                // ищем строку с наибольшим по модулю элементом в столбце k
                int maxRow = k;
                for (int i = k + 1; i < n; i++)
                {
                    if (Math.Abs(a[i, k]) > Math.Abs(a[maxRow, k]))
                        maxRow = i;
                }

                if (Math.Abs(a[maxRow, k]) < 1e-12)
                    return 0;

                // при обмене строк определитель меняет знак
                if (maxRow != k)
                {
                    for (int j = 0; j < n; j++)
                    {
                        double temp = a[k, j];
                        a[k, j] = a[maxRow, j];
                        a[maxRow, j] = temp;
                    }
                    det = -det;
                }

                det = det * a[k, k];

                // обнуляем элементы под главным
                for (int i = k + 1; i < n; i++)
                {
                    double m = a[i, k] / a[k, k];
                    for (int j = k; j < n; j++)
                        a[i, j] = a[i, j] - m * a[k, j];
                }
            }

            return det;
        }

        // вывод результата в любую таблицу
        private void ShowResult(DataGridView grid, double[] x)
        {
            grid.Columns.Clear();

            grid.Columns.Add("colX", "x");
            grid.Columns[0].Width = 100;
            grid.Columns[0].DefaultCellStyle.Format = "0.#####";
            grid.Rows.Add(x.Length);

            for (int i = 0; i < x.Length; i++)
            {
                grid.Rows[i].HeaderCell.Value = "x" + (i + 1);
                grid.Rows[i].Cells[0].Value = Math.Round(x[i], 5) + 0.0;
            }
        }

        // ОЧИСТИТЬ / СЛУЧАЙНЫЕ

        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isBusy) return;

            // обнуляем входные таблицы
            int n = (int)numericUpDown1.Value;

            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                    InputAdataGridView.Rows[i].Cells[j].Value = 0;

                InputBdataGridView.Rows[i].Cells[0].Value = 0;
            }

            // очищаем результаты и подписи
            GaussdataGridView.Columns.Clear();
            Gausslabel.Text = "";

            GaussJordandataGridView.Columns.Clear();
            GaussJordanlabel.Text = "";

            CramerdataGridView.Columns.Clear();
            Cramerlabel.Text = "";

            // сбрасываем данные для экспорта
            lastA = null;
            lastB = null;
            lastGauss = null;
            lastGaussJordan = null;
            lastCramer = null;
        }

        private void случайныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isBusy) return;

            int n = (int)numericUpDown1.Value;

            // таблицы должны соответствовать выбранному размеру
            if (InputAdataGridView.RowCount != n)
                CreateTables();

            double[,] A = new double[n, n];

            // генерируем A, пока определитель не станет ненулевым
            do
            {
                for (int i = 0; i < n; i++)
                    for (int j = 0; j < n; j++)
                        A[i, j] = rnd.Next(-10, 11);   // от -10 до 10
            }
            while (Math.Abs(Determinant(A)) < 1e-12);

            // записываем A в таблицу
            for (int i = 0; i < n; i++)
                for (int j = 0; j < n; j++)
                    InputAdataGridView.Rows[i].Cells[j].Value = (int)A[i, j];

            // генерируем и записываем B
            for (int i = 0; i < n; i++)
                InputBdataGridView.Rows[i].Cells[0].Value = rnd.Next(-10, 11);
        }

        // ЗАГРУЗКА ДАННЫХ

        // Заполняет таблицы по готовым массивам
        private void FillTables(int n, int[,] A, int[] B)
        {
            numericUpDown1.Value = n;
            if (InputAdataGridView.RowCount != n)
                CreateTables();

            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                    InputAdataGridView.Rows[i].Cells[j].Value = A[i, j];

                InputBdataGridView.Rows[i].Cells[0].Value = B[i];
            }
        }

        // Показывает окошко с полем ввода. Возвращает текст или null, если нажали «Отмена»
        private string AskForText(string title, string prompt)
        {
            Form dlg = new Form();
            dlg.Text = title;
            dlg.ClientSize = new Size(460, 120);
            dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
            dlg.StartPosition = FormStartPosition.CenterParent;
            dlg.MaximizeBox = false;
            dlg.MinimizeBox = false;

            Label label = new Label();
            label.Text = prompt;
            label.Left = 12;
            label.Top = 12;
            label.AutoSize = true;

            TextBox box = new TextBox();
            box.Left = 12;
            box.Top = 36;
            box.Width = 436;

            Button ok = new Button();
            ok.Text = "Загрузить";
            ok.Left = 258;
            ok.Top = 76;
            ok.Width = 90;
            ok.DialogResult = DialogResult.OK;

            Button cancel = new Button();
            cancel.Text = "Отмена";
            cancel.Left = 358;
            cancel.Top = 76;
            cancel.Width = 90;
            cancel.DialogResult = DialogResult.Cancel;

            dlg.Controls.Add(label);
            dlg.Controls.Add(box);
            dlg.Controls.Add(ok);
            dlg.Controls.Add(cancel);
            dlg.AcceptButton = ok;
            dlg.CancelButton = cancel;

            string result = null;
            if (dlg.ShowDialog(this) == DialogResult.OK)
                result = box.Text;

            dlg.Dispose();
            return result;
        }

        // Загрузка из файла Excel (.xlsx)
        private void загрузитьИзФайлаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isBusy) return;

            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Файлы Excel (*.xlsx)|*.xlsx";

            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                using (XLWorkbook book = new XLWorkbook(dlg.FileName))
                {
                    IXLWorksheet sheet = book.Worksheet(1);   // первый лист

                    int rows = sheet.LastRowUsed().RowNumber();
                    int cols = sheet.LastColumnUsed().ColumnNumber();
                    int n = rows;

                    if (n < 2 || n > 50)
                    {
                        MessageBox.Show("Размер системы должен быть от 2 до 50. В файле строк: " + rows,
                                        "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (cols != n + 1)
                    {
                        MessageBox.Show("В файле должно быть " + n + " строк и " + (n + 1) +
                                        " столбцов (матрица A и столбец B). Сейчас столбцов: " + cols,
                                        "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    int[,] A = new int[n, n];
                    int[] B = new int[n];

                    for (int i = 0; i < n; i++)
                    {
                        for (int j = 0; j <= n; j++)
                        {
                            string s = sheet.Cell(i + 1, j + 1).GetString().Trim();
                            int value;

                            if (!int.TryParse(s, out value))
                            {
                                MessageBox.Show("В ячейке (строка " + (i + 1) + ", столбец " + (j + 1) +
                                                ") должно быть целое число!",
                                                "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            if (j < n)
                                A[i, j] = value;
                            else
                                B[i] = value;
                        }
                    }

                    FillTables(n, A, B);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось прочитать файл (возможно, он открыт в Excel):\n" + ex.Message,
                                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Загрузка из Google Таблиц по ссылке
        // (таблица должна быть открыта для всех, у кого есть ссылка)
        private async void загрузитьПоСсылкеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isBusy) return;

            // 1. Спрашиваем ссылку
            string link = AskForText("Загрузка из Google Sheets", "Вставьте ссылку на таблицу Google Sheets:");

            if (string.IsNullOrWhiteSpace(link)) return;

            link = link.Trim();

            // 2. Достаём из ссылки id таблицы и номер листа (gid)
            Match idMatch = Regex.Match(link, @"/spreadsheets/d/([a-zA-Z0-9_-]+)");
            if (!idMatch.Success)
            {
                MessageBox.Show("Это не похоже на ссылку на Google Sheets!",
                                "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string id = idMatch.Groups[1].Value;

            // Если в ссылке есть номер листа (gid) - берём его, иначе Google отдаст первый лист
            string gidPart = "";
            Match gidMatch = Regex.Match(link, @"[#&?]gid=([0-9]+)");
            if (gidMatch.Success)
            {
                gidPart = "&gid=" + gidMatch.Groups[1].Value;
            }

            // ссылка для скачивания листа в формате CSV
            string csvUrl = "https://docs.google.com/spreadsheets/d/" + id + "/export?format=csv" + gidPart;

            // 3. Скачиваем
            string text = "";
            isBusy = true; // пока идёт загрузка, другие кнопки не работают

            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                using (HttpClient client = new HttpClient())
                {
                    text = await client.GetStringAsync(csvUrl);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось скачать таблицу.\n" +
                                "Проверьте интернет и то, что доступ к таблице открыт по ссылке.\n\n" + ex.Message,
                                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                isBusy = false;
            }

            // если таблица закрыта, Google присылает страницу входа вместо данных
            if (text.Contains("<html") || text.Contains("<!DOCTYPE"))
            {
                MessageBox.Show("Таблица закрыта. Откройте доступ: Настройки доступа -> \"Все, у кого есть ссылка\" (читатель).",
                                "Нет доступа", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 4. Разбиваем на строки, пустые строки пропускаем
            string[] allLines = text.Split('\n');
            List<string> lines = new List<string>();
            foreach (string line in allLines)
            {
                if (line.Trim(',', ' ', '\r', '"') != "")
                    lines.Add(line.Trim('\r'));
            }

            int n = lines.Count;

            if (n < 2 || n > 50)
            {
                MessageBox.Show("Размер системы должен быть от 2 до 50. В таблице строк: " + n,
                                "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int[,] A = new int[n, n];
            int[] B = new int[n];

            for (int i = 0; i < n; i++)
            {
                string[] cells = lines[i].Split(',');

                if (cells.Length < n + 1)
                {
                    MessageBox.Show("В строке " + (i + 1) + " должно быть " + (n + 1) + " чисел.",
                                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                for (int j = 0; j <= n; j++)
                {
                    string s = cells[j].Trim().Trim('"');
                    int value;

                    if (!int.TryParse(s, out value))
                    {
                        MessageBox.Show("В ячейке (строка " + (i + 1) + ", столбец " + (j + 1) +
                                        ") должно быть целое число!",
                                        "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    if (j < n)
                        A[i, j] = value;
                    else
                        B[i] = value;
                }
            }

            // 5. Ошибок нет, выводим в таблицы
            FillTables(n, A, B);
        }

        // ЭКСПОРТ

        // Кнопка "Скачать": сохраняет A, B и X на НОВЫЙ лист Excel-файла
        private void скачатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isBusy) return;

            if (lastA == null || (lastGauss == null && lastGaussJordan == null && lastCramer == null))
            {
                MessageBox.Show("Сначала выполните расчёт (кнопка \"Рассчитать\")!",
                                "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Title = "Экспорт системы и решения";
            dialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx";
            dialog.FileName = "Система уравнений.xlsx";
            dialog.OverwritePrompt = false;   // выбор существующего файла не спрашивает о перезаписи

            if (dialog.ShowDialog() != DialogResult.OK) return;

            try
            {
                int n = lastB.Length;

                // если файл есть, открываем его и добавляем лист, иначе создаём новый
                XLWorkbook workbook;
                if (File.Exists(dialog.FileName))
                    workbook = new XLWorkbook(dialog.FileName);
                else
                    workbook = new XLWorkbook();

                using (workbook)
                {
                    // имя нового листа не должно совпадать с уже существующими
                    string sheetName = "Результат " + DateTime.Now.ToString("dd.MM HH-mm-ss");
                    IXLWorksheet sheet = workbook.Worksheets.Add(sheetName);

                    // Матрица A
                    sheet.Cell(1, 1).Value = "Матрица A";
                    for (int j = 0; j < n; j++)
                        sheet.Cell(2, j + 1).Value = "x" + (j + 1);

                    for (int i = 0; i < n; i++)
                        for (int j = 0; j < n; j++)
                            sheet.Cell(i + 3, j + 1).Value = lastA[i, j];

                    // Вектор B (через один пустой столбец)
                    int colB = n + 2;
                    sheet.Cell(1, colB).Value = "Вектор B";
                    sheet.Cell(2, colB).Value = "B";
                    for (int i = 0; i < n; i++)
                        sheet.Cell(i + 3, colB).Value = lastB[i];

                    // Векторы X
                    int col = colB + 2;

                    if (lastGauss != null)
                    {
                        sheet.Cell(1, col).Value = "Вектор X";
                        sheet.Cell(2, col).Value = "Гаусс";
                        for (int i = 0; i < n; i++)
                            sheet.Cell(i + 3, col).Value = Math.Round(lastGauss[i], 5) + 0.0;
                        col++;
                    }

                    if (lastGaussJordan != null)
                    {
                        sheet.Cell(1, col).Value = "Вектор X";
                        sheet.Cell(2, col).Value = "Гаусс-Жордан";
                        for (int i = 0; i < n; i++)
                            sheet.Cell(i + 3, col).Value = Math.Round(lastGaussJordan[i], 5) + 0.0;
                        col++;
                    }

                    if (lastCramer != null)
                    {
                        sheet.Cell(1, col).Value = "Вектор X";
                        sheet.Cell(2, col).Value = "Крамер";
                        for (int i = 0; i < n; i++)
                            sheet.Cell(i + 3, col).Value = Math.Round(lastCramer[i], 5) + 0.0;
                        col++;
                    }

                    // оформление: жирные заголовки, ширина по содержимому
                    sheet.Range(1, 1, 2, col - 1).Style.Font.Bold = true;
                    sheet.Columns().AdjustToContents();

                    workbook.SaveAs(dialog.FileName);
                }

                MessageBox.Show("Файл сохранён!",
                                "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось сохранить файл (возможно, он открыт):\n" + ex.Message,
                                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}