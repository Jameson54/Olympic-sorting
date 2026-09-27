using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace Lab4SortingApp;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

internal sealed class MainForm : Form
{
    private readonly DataGridView inputGrid = new();
    private readonly FlowLayoutPanel algorithmBar = new();
    private readonly FlowLayoutPanel charts = new();
    private readonly ToolStripStatusLabel status = new("Готово");
    private readonly Dictionary<string, CheckBox> checks = new();
    private readonly Dictionary<string, SortView> views = new();
    private readonly RadioButton ascending = new() { Text = "По возрастанию", Checked = true, AutoSize = true };
    private readonly RadioButton descending = new() { Text = "По убыванию", AutoSize = true };
    private CancellationTokenSource? cancellation;

    private static readonly string[] AlgorithmNames =
    {
        "Пузырьковая", "Вставками", "Шейкерная", "Быстрая", "BOGO"
    };

    public MainForm()
    {
        Text = "Лабораторная работа №4 — Олимпиадные сортировки";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(1050, 700);
        Width = 1280;
        Height = 820;

        var menu = BuildMenu();
        MainMenuStrip = menu;
        Controls.Add(menu);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            RowCount = 3,
            ColumnCount = 1,
            Padding = new Padding(8),
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 190));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 55));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        inputGrid.Dock = DockStyle.Fill;
        inputGrid.AllowUserToAddRows = true;
        inputGrid.AllowUserToDeleteRows = true;
        inputGrid.RowHeadersVisible = false;
        inputGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        inputGrid.Columns.Add("Value", "Число");
        root.Controls.Add(inputGrid, 0, 0);

        algorithmBar.Dock = DockStyle.Fill;
        algorithmBar.FlowDirection = FlowDirection.LeftToRight;
        algorithmBar.WrapContents = true;
        algorithmBar.Padding = new Padding(4, 10, 4, 4);
        foreach (var name in AlgorithmNames)
        {
            var check = new CheckBox { Text = name, AutoSize = true, Margin = new Padding(8, 3, 8, 3) };
            checks[name] = check;
            algorithmBar.Controls.Add(check);
        }
        algorithmBar.Controls.Add(new Label { Text = "     Направление:", AutoSize = true, Margin = new Padding(12, 5, 2, 2) });
        algorithmBar.Controls.Add(ascending);
        algorithmBar.Controls.Add(descending);
        root.Controls.Add(algorithmBar, 0, 1);

        charts.Dock = DockStyle.Fill;
        charts.AutoScroll = true;
        charts.WrapContents = true;
        charts.FlowDirection = FlowDirection.LeftToRight;
        root.Controls.Add(charts, 0, 2);

        var statusStrip = new StatusStrip();
        statusStrip.Items.Add(status);
        Controls.Add(statusStrip);
        Controls.Add(root);
        root.BringToFront();

        GenerateData();
    }

    private MenuStrip BuildMenu()
    {
        var menu = new MenuStrip();

        var data = new ToolStripMenuItem("Данные");
        data.DropDownItems.Add("Сгенерировать", null, (_, _) => GenerateData());
        data.DropDownItems.Add("Загрузить CSV / TSV", null, async (_, _) => await LoadDelimitedFileAsync());
        data.DropDownItems.Add("Загрузить Excel XLSX", null, (_, _) => LoadXlsx());
        data.DropDownItems.Add("Загрузить Google Таблицу", null, async (_, _) => await LoadGoogleSheetAsync());
        data.DropDownItems.Add(new ToolStripSeparator());
        data.DropDownItems.Add("Очистить", null, (_, _) => ClearAll());

        var calculation = new ToolStripMenuItem("Расчёт");
        calculation.DropDownItems.Add("Запустить", null, async (_, _) => await RunSelectedAsync());
        calculation.DropDownItems.Add("Остановить", null, (_, _) => cancellation?.Cancel());

        menu.Items.Add(data);
        menu.Items.Add(calculation);
        return menu;
    }

    private void GenerateData()
    {
        var countText = SimpleDialog.Show("Количество элементов (2–100):", "Генерация", "20");
        if (countText is null) return;
        if (!int.TryParse(countText, out var count) || count is < 2 or > 100)
        {
            Error("Введите целое число от 2 до 100.");
            return;
        }

        var random = new Random();
        SetInput(Enumerable.Range(0, count).Select(_ => random.Next(-100, 101)));
        status.Text = $"Сгенерировано элементов: {count}";
    }

    private void ClearAll()
    {
        cancellation?.Cancel();
        inputGrid.Rows.Clear();
        charts.Controls.Clear();
        views.Clear();
        status.Text = "Очищено";
    }

    private async Task LoadDelimitedFileAsync()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Табличные файлы (*.csv;*.tsv;*.txt)|*.csv;*.tsv;*.txt|Все файлы (*.*)|*.*"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        try
        {
            var text = await File.ReadAllTextAsync(dialog.FileName);
            SetInput(ParseNumbers(text));
            status.Text = "Данные загружены";
        }
        catch (Exception ex) { Error(ex.Message); }
    }

    private void LoadXlsx()
    {
        using var dialog = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        try
        {
            SetInput(XlsxReader.ReadNumbers(dialog.FileName));
            status.Text = "Данные из Excel загружены";
        }
        catch (Exception ex) { Error("Не удалось прочитать XLSX: " + ex.Message); }
    }

    private async Task LoadGoogleSheetAsync()
    {
        var source = SimpleDialog.Show(
            "Вставьте общедоступную ссылку Google Таблицы или ссылку публикации CSV:",
            "Google Таблица", "");
        if (string.IsNullOrWhiteSpace(source)) return;

        try
        {
            var url = GoogleCsvUrl(source);
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
            var csv = await client.GetStringAsync(url);
            SetInput(ParseNumbers(csv));
            status.Text = "Данные из Google Таблицы загружены";
        }
        catch (Exception ex)
        {
            Error("Проверьте, что таблица доступна по ссылке. " + ex.Message);
        }
    }

    private static string GoogleCsvUrl(string source)
    {
        source = source.Trim();
        if (source.Contains("output=csv", StringComparison.OrdinalIgnoreCase) ||
            source.Contains("format=csv", StringComparison.OrdinalIgnoreCase)) return source;

        var marker = "/spreadsheets/d/";
        var p = source.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (p < 0) throw new FormatException("Это не ссылка Google Таблицы.");
        var start = p + marker.Length;
        var end = source.IndexOf('/', start);
        var id = end < 0 ? source[start..] : source[start..end];
        var gid = "0";
        var gidPos = source.IndexOf("gid=", StringComparison.OrdinalIgnoreCase);
        if (gidPos >= 0)
        {
            gid = new string(source[(gidPos + 4)..].TakeWhile(char.IsDigit).ToArray());
            if (gid.Length == 0) gid = "0";
        }
        return $"https://docs.google.com/spreadsheets/d/{id}/export?format=csv&gid={gid}";
    }

    private static IEnumerable<int> ParseNumbers(string text)
    {
        var tokens = text.Split(new[] { ',', ';', '\t', '\r', '\n', ' ' }, StringSplitOptions.RemoveEmptyEntries);
        var values = new List<int>();
        foreach (var token in tokens)
        {
            var normalized = token.Trim().Trim('"');
            if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ||
                int.TryParse(normalized, NumberStyles.Integer, CultureInfo.CurrentCulture, out value))
                values.Add(value);
        }
        if (values.Count == 0) throw new FormatException("В файле не найдено целых чисел.");
        return values;
    }

  private void SetInput(IEnumerable<int> source)
  {
    var values = source.Take(500).ToArray();

    if (values.Length == 0)
      throw new InvalidOperationException("Нет данных для загрузки.");

    inputGrid.Rows.Clear();

    foreach (var value in values)
      inputGrid.Rows.Add(new object[] { value });
  }

  private int[] ReadInput()
    {
        var values = new List<int>();
        for (var i = 0; i < inputGrid.Rows.Count; i++)
        {
            if (inputGrid.Rows[i].IsNewRow) continue;
            var raw = Convert.ToString(inputGrid.Rows[i].Cells[0].Value)?.Trim();
            if (string.IsNullOrEmpty(raw)) continue;
            if (!int.TryParse(raw, out var value))
                throw new FormatException($"Строка {i + 1}: «{raw}» не является целым числом.");
            values.Add(value);
        }
        if (values.Count < 2) throw new InvalidOperationException("Введите минимум два целых числа.");
        if (values.Count > 500) throw new InvalidOperationException("Допустимо не более 500 элементов.");
        return values.ToArray();
    }

    private async Task RunSelectedAsync()
    {
        if (cancellation is not null)
        {
            Error("Сортировка уже выполняется.");
            return;
        }

        try
        {
            var input = ReadInput();
            var selected = AlgorithmNames.Where(name => checks[name].Checked).ToArray();
            if (selected.Length == 0) throw new InvalidOperationException("Выберите хотя бы один алгоритм.");
            if (selected.Contains("BOGO") && input.Length > 9)
                throw new InvalidOperationException("Для BOGO используйте не более 9 элементов.");

            cancellation = new CancellationTokenSource();
            var token = cancellation.Token;
            var direction = ascending.Checked ? 1 : -1;
            charts.Controls.Clear();
            views.Clear();

            foreach (var name in selected)
            {
                var view = new SortView(name, input);
                views[name] = view;
                charts.Controls.Add(view);
            }

            status.Text = "Выполняется сортировка…";
            var tasks = selected.Select(name => AnimateAsync(name, input, direction, token)).ToArray();
            await Task.WhenAll(tasks);
            status.Text = "Готово. Самый быстрый: " + views.Values
                .Where(v => v.Elapsed is not null)
                .OrderBy(v => v.Elapsed)
                .First().Title;
        }
        catch (OperationCanceledException) { status.Text = "Остановлено"; }
        catch (Exception ex) { Error(ex.Message); }
        finally
        {
            cancellation?.Dispose();
            cancellation = null;
        }
    }

    private async Task AnimateAsync(string name, int[] input, int direction, CancellationToken token)
    {
        var view = views[name];
        var frames = SortAlgorithms.Frames(name, input, direction).GetEnumerator();
        var displayed = 0;
        while (frames.MoveNext())
        {
            token.ThrowIfCancellationRequested();
            view.SetValues(frames.Current);
            displayed++;
            if (displayed >= 1500) break;
            await Task.Delay(input.Length <= 30 ? 25 : 5, token);
        }

        var copy = (int[])input.Clone();
        var watch = Stopwatch.StartNew();
        var completed = SortAlgorithms.Sort(name, copy, direction, token);
        watch.Stop();
        view.SetValues(copy);
        view.SetResult(completed ? watch.Elapsed : null,
            completed ? $"{watch.Elapsed.TotalMilliseconds:F4} мс" : "Лимит BOGO исчерпан");
    }

    private void Error(string message)
    {
        status.Text = "Ошибка";
        MessageBox.Show(this, message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}

internal static class SortAlgorithms
{
    private static bool Before(int left, int right, int direction) => direction > 0 ? left < right : left > right;
    private static bool After(int left, int right, int direction) => direction > 0 ? left > right : left < right;

    public static IEnumerable<int[]> Frames(string name, int[] source, int direction)
    {
        var a = (int[])source.Clone();
        return name switch
        {
            "Пузырьковая" => BubbleFrames(a, direction),
            "Вставками" => InsertionFrames(a, direction),
            "Шейкерная" => CocktailFrames(a, direction),
            "Быстрая" => QuickFrames(a, 0, a.Length - 1, direction),
            "BOGO" => BogoFrames(a, direction),
            _ => throw new ArgumentOutOfRangeException(nameof(name))
        };
    }

    public static bool Sort(string name, int[] a, int direction, CancellationToken token)
    {
        switch (name)
        {
            case "Пузырьковая": Bubble(a, direction); return true;
            case "Вставками": Insertion(a, direction); return true;
            case "Шейкерная": Cocktail(a, direction); return true;
            case "Быстрая": Quick(a, 0, a.Length - 1, direction); return true;
            case "BOGO": return Bogo(a, direction, token);
            default: throw new ArgumentOutOfRangeException(nameof(name));
        }
    }

    private static IEnumerable<int[]> BubbleFrames(int[] a, int d)
    {
        yield return (int[])a.Clone();
        for (var end = a.Length - 1; end > 0; end--)
            for (var i = 0; i < end; i++)
                if (After(a[i], a[i + 1], d)) { (a[i], a[i + 1]) = (a[i + 1], a[i]); yield return (int[])a.Clone(); }
    }

    private static IEnumerable<int[]> InsertionFrames(int[] a, int d)
    {
        yield return (int[])a.Clone();
        for (var i = 1; i < a.Length; i++)
        {
            var x = a[i]; var j = i - 1;
            while (j >= 0 && After(a[j], x, d)) { a[j + 1] = a[j--]; yield return (int[])a.Clone(); }
            a[j + 1] = x; yield return (int[])a.Clone();
        }
    }

    private static IEnumerable<int[]> CocktailFrames(int[] a, int d)
    {
        yield return (int[])a.Clone();
        var left = 0; var right = a.Length - 1;
        while (left < right)
        {
            for (var i = left; i < right; i++)
                if (After(a[i], a[i + 1], d)) { (a[i], a[i + 1]) = (a[i + 1], a[i]); yield return (int[])a.Clone(); }
            right--;
            for (var i = right; i > left; i--)
                if (After(a[i - 1], a[i], d)) { (a[i - 1], a[i]) = (a[i], a[i - 1]); yield return (int[])a.Clone(); }
            left++;
        }
    }

    private static IEnumerable<int[]> QuickFrames(int[] a, int left, int right, int d)
    {
        if (left >= right) yield break;
        var i = left; var j = right; var pivot = a[(left + right) / 2];
        while (i <= j)
        {
            while (Before(a[i], pivot, d)) i++;
            while (After(a[j], pivot, d)) j--;
            if (i <= j) { (a[i], a[j]) = (a[j], a[i]); i++; j--; yield return (int[])a.Clone(); }
        }
        foreach (var frame in QuickFrames(a, left, j, d)) yield return frame;
        foreach (var frame in QuickFrames(a, i, right, d)) yield return frame;
    }

    private static IEnumerable<int[]> BogoFrames(int[] a, int d)
    {
        yield return (int[])a.Clone();
        var random = new Random();
        for (var attempt = 0; attempt < 1500 && !IsSorted(a, d); attempt++)
        {
            random.Shuffle(a);
            yield return (int[])a.Clone();
        }
    }

    private static void Bubble(int[] a, int d)
    {
        for (var end = a.Length - 1; end > 0; end--)
            for (var i = 0; i < end; i++) if (After(a[i], a[i + 1], d)) (a[i], a[i + 1]) = (a[i + 1], a[i]);
    }

    private static void Insertion(int[] a, int d)
    {
        for (var i = 1; i < a.Length; i++)
        {
            var x = a[i]; var j = i - 1;
            while (j >= 0 && After(a[j], x, d)) { a[j + 1] = a[j--]; }
            a[j + 1] = x;
        }
    }

    private static void Cocktail(int[] a, int d)
    {
        var left = 0; var right = a.Length - 1;
        while (left < right)
        {
            for (var i = left; i < right; i++) if (After(a[i], a[i + 1], d)) (a[i], a[i + 1]) = (a[i + 1], a[i]);
            right--;
            for (var i = right; i > left; i--) if (After(a[i - 1], a[i], d)) (a[i - 1], a[i]) = (a[i], a[i - 1]);
            left++;
        }
    }

    private static void Quick(int[] a, int left, int right, int d)
    {
        if (left >= right) return;
        var i = left; var j = right; var pivot = a[(left + right) / 2];
        while (i <= j)
        {
            while (Before(a[i], pivot, d)) i++;
            while (After(a[j], pivot, d)) j--;
            if (i <= j) { (a[i], a[j]) = (a[j], a[i]); i++; j--; }
        }
        if (left < j) Quick(a, left, j, d);
        if (i < right) Quick(a, i, right, d);
    }

    private static bool Bogo(int[] a, int d, CancellationToken token)
    {
        var random = new Random();
        for (var attempt = 0; attempt < 100_000; attempt++)
        {
            token.ThrowIfCancellationRequested();
            if (IsSorted(a, d)) return true;
            random.Shuffle(a);
        }
        return IsSorted(a, d);
    }

    private static bool IsSorted(int[] a, int d)
    {
        for (var i = 1; i < a.Length; i++) if (After(a[i - 1], a[i], d)) return false;
        return true;
    }

    private static void Shuffle(this Random random, int[] a)
    {
        for (var i = a.Length - 1; i > 0; i--)
        {
            var j = random.Next(i + 1);
            (a[i], a[j]) = (a[j], a[i]);
        }
    }
}

internal sealed class SortView : UserControl
{
    private int[] values;
    private readonly Label caption = new();
    public string Title { get; }
    public TimeSpan? Elapsed { get; private set; }

    public SortView(string title, int[] source)
    {
        Title = title;
        values = (int[])source.Clone();
        Width = 390;
        Height = 250;
        Margin = new Padding(7);
        BorderStyle = BorderStyle.FixedSingle;
        DoubleBuffered = true;
        caption.Text = title;
        caption.Dock = DockStyle.Top;
        caption.Height = 30;
        caption.TextAlign = ContentAlignment.MiddleCenter;
        caption.Font = new Font(Font, FontStyle.Bold);
        Controls.Add(caption);
    }

    public void SetValues(int[] newValues) { values = (int[])newValues.Clone(); Invalidate(); }
    public void SetResult(TimeSpan? elapsed, string text) { Elapsed = elapsed; caption.Text = $"{Title}: {text}"; }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        if (values.Length == 0) return;
        var area = new Rectangle(8, caption.Bottom + 8, ClientSize.Width - 16, ClientSize.Height - caption.Bottom - 16);
        var min = values.Min(); var max = values.Max(); var range = Math.Max(1, max - min);
        var width = Math.Max(1f, (float)area.Width / values.Length);
        using var brush = new SolidBrush(Color.SteelBlue);
        using var axis = new Pen(Color.LightGray);
        e.Graphics.DrawRectangle(axis, area);
        for (var i = 0; i < values.Length; i++)
        {
            var height = Math.Max(2f, (values[i] - min + 1f) / (range + 1f) * (area.Height - 4));
            e.Graphics.FillRectangle(brush, area.Left + i * width + 1, area.Bottom - height - 1, Math.Max(1, width - 2), height);
        }
    }
}

internal static class XlsxReader
{
    public static IEnumerable<int> ReadNumbers(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var shared = ReadSharedStrings(archive);
        var sheet = archive.GetEntry("xl/worksheets/sheet1.xml")
                    ?? throw new InvalidDataException("Первый лист не найден.");
        using var stream = sheet.Open();
        var doc = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var values = new List<int>();
        foreach (var cell in doc.Descendants(ns + "c"))
        {
            var type = (string?)cell.Attribute("t");
            var raw = cell.Element(ns + "v")?.Value ?? cell.Descendants(ns + "t").FirstOrDefault()?.Value;
            if (raw is null) continue;
            if (type == "s" && int.TryParse(raw, out var index) && index >= 0 && index < shared.Count) raw = shared[index];
            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) &&
                number >= int.MinValue && number <= int.MaxValue && Math.Abs(number % 1) < 0.0000001)
                values.Add((int)number);
        }
        if (values.Count == 0) throw new InvalidDataException("На первом листе не найдено целых чисел.");
        return values;
    }

    private static List<string> ReadSharedStrings(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry is null) return new List<string>();
        using var stream = entry.Open();
        var doc = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        return doc.Descendants(ns + "si").Select(si => string.Concat(si.Descendants(ns + "t").Select(t => t.Value))).ToList();
    }
}

internal static class SimpleDialog
{
    public static string? Show(string prompt, string title, string initial)
    {
        using var form = new Form
        {
            Text = title, StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog, MinimizeBox = false, MaximizeBox = false,
            Width = 570, Height = 175
        };
        var label = new Label { Text = prompt, Left = 12, Top = 12, Width = 530, Height = 36 };
        var box = new TextBox { Text = initial, Left = 12, Top = 52, Width = 530 };
        var ok = new Button { Text = "ОК", DialogResult = DialogResult.OK, Left = 366, Width = 85, Top = 88 };
        var cancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, Left = 457, Width = 85, Top = 88 };
        form.Controls.AddRange(new Control[] { label, box, ok, cancel });
        form.AcceptButton = ok; form.CancelButton = cancel;
        return form.ShowDialog() == DialogResult.OK ? box.Text : null;
    }
}
