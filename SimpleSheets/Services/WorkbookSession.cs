using System.Collections;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using ClosedXML.Excel;

namespace SimpleSheets.Services;

public enum CellHorizontalAlign
{
    Left,
    Center,
    Right
}

public readonly record struct CellDisplay(
    string Text,
    double FontSize,
    bool Bold,
    bool Italic,
    bool WrapText,
    CellHorizontalAlign Align,
    byte? Red,
    byte? Green,
    byte? Blue);

public sealed class WorkbookSession : IDisposable
{
    const int MinimumRows = 50;
    const int MinimumColumns = 26;

    static readonly Dictionary<string, string> RussianFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        ["СУММ"] = "SUM",
        ["СУММЕСЛИ"] = "SUMIF",
        ["СУММЕСЛИМН"] = "SUMIFS",
        ["СРЗНАЧ"] = "AVERAGE",
        ["СРЗНАЧЕСЛИ"] = "AVERAGEIF",
        ["СРЗНАЧЕСЛИМН"] = "AVERAGEIFS",
        ["ЕСЛИ"] = "IF",
        ["ЕСЛИОШИБКА"] = "IFERROR",
        ["СЧЁТ"] = "COUNT",
        ["СЧЕТ"] = "COUNT",
        ["СЧЁТЗ"] = "COUNTA",
        ["СЧЕТЗ"] = "COUNTA",
        ["СЧЁТЕСЛИ"] = "COUNTIF",
        ["СЧЕТЕСЛИ"] = "COUNTIF",
        ["СЧЁТЕСЛИМН"] = "COUNTIFS",
        ["МАКС"] = "MAX",
        ["МИН"] = "MIN",
        ["И"] = "AND",
        ["ИЛИ"] = "OR",
        ["НЕ"] = "NOT",
        ["ОКРУГЛ"] = "ROUND",
        ["ОКРУГЛВВЕРХ"] = "ROUNDUP",
        ["ОКРУГЛВНИЗ"] = "ROUNDDOWN",
        ["ЦЕЛОЕ"] = "INT",
        ["ОСТАТ"] = "MOD",
        ["СТЕПЕНЬ"] = "POWER",
        ["КОРЕНЬ"] = "SQRT",
        ["СЕГОДНЯ"] = "TODAY",
        ["ТДАТА"] = "NOW",
        ["ДАТА"] = "DATE",
        ["ДЕНЬ"] = "DAY",
        ["МЕСЯЦ"] = "MONTH",
        ["ГОД"] = "YEAR",
        ["СЦЕПИТЬ"] = "CONCATENATE",
        ["СЦЕП"] = "CONCAT",
        ["ЛЕВСИМВ"] = "LEFT",
        ["ПРАВСИМВ"] = "RIGHT",
        ["ПСТР"] = "MID",
        ["ДЛСТР"] = "LEN",
        ["ПРОПИСН"] = "UPPER",
        ["СТРОЧН"] = "LOWER",
        ["СЖПРОБЕЛЫ"] = "TRIM",
        ["ВПР"] = "VLOOKUP",
        ["ГПР"] = "HLOOKUP",
        ["ИНДЕКС"] = "INDEX",
        ["ПОИСКПОЗ"] = "MATCH",
        ["ИСТИНА"] = "TRUE",
        ["ЛОЖЬ"] = "FALSE",
        ["ПИ"] = "PI",
        ["ЗНАК"] = "SIGN",
        ["ABS"] = "ABS"
    };

    readonly XLWorkbook _workbook;
    readonly List<SheetState> _sheets;
    string? _path;

    WorkbookSession(XLWorkbook workbook, string? path)
    {
        _workbook = workbook;
        _path = path;
        _sheets = workbook.Worksheets.Select(LoadState).ToList();
        if (_sheets.Count == 0)
            throw new InvalidDataException("В книге нет листов.");
    }

    public string? FilePath => _path;

    public bool IsDirty { get; private set; }

    public string Title => _path == null ? "Книга1" : Path.GetFileName(_path);

    public int ActiveSheetIndex { get; private set; }

    public IReadOnlyList<string> SheetNames => _workbook.Worksheets.Select(sheet => sheet.Name).ToList();

    public int DisplayRows => State.DisplayRows;

    public int DisplayColumns => State.DisplayColumns;

    public double DefaultColumnWidth => State.DefaultColumnWidth;

    public double DefaultRowHeight => State.DefaultRowHeight;

    public int LastContentRow => Active.LastRowUsed(XLCellsUsedOptions.Contents)?.RowNumber() ?? 0;

    public int LastContentColumn => Active.LastColumnUsed(XLCellsUsedOptions.Contents)?.ColumnNumber() ?? 0;

    SheetState State => _sheets[ActiveSheetIndex];

    IXLWorksheet Active => _workbook.Worksheet(ActiveSheetIndex + 1);

    public static WorkbookSession CreateNew()
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Лист1");
        sheet.Style.Font.FontName = AppFonts.ExcelName;
        sheet.Style.Font.FontSize = 11;
        return new WorkbookSession(workbook, null);
    }

    public static WorkbookSession Open(string path)
    {
        var workbook = new XLWorkbook(path);
        return new WorkbookSession(workbook, path);
    }

    public void SetActiveSheet(int index)
    {
        if (index < 0 || index >= _sheets.Count)
            throw new ArgumentOutOfRangeException(nameof(index));
        ActiveSheetIndex = index;
    }

    public bool EnsureCapacity(int row, int column)
    {
        var state = State;
        var changed = false;
        if (row >= state.DisplayRows)
        {
            state.DisplayRows = row + 20;
            changed = true;
        }

        if (column >= state.DisplayColumns)
        {
            state.DisplayColumns = column + 5;
            changed = true;
        }

        return changed;
    }

    public double GetColumnWidth(int column)
    {
        if (column < 1)
            throw new ArgumentOutOfRangeException(nameof(column));
        return State.ColumnWidths.TryGetValue(column, out var width)
            ? width
            : State.DefaultColumnWidth;
    }

    public double GetRowHeight(int row)
    {
        if (row < 1)
            throw new ArgumentOutOfRangeException(nameof(row));
        return State.RowHeights.TryGetValue(row, out var height)
            ? height
            : State.DefaultRowHeight;
    }

    public void SetColumnWidth(int column, double width)
    {
        width = Math.Clamp(width, 0, 255);
        Active.Column(column).Width = width;
        State.ColumnWidths[column] = width;
        IsDirty = true;
    }

    public void SetRowHeight(int row, double points)
    {
        points = Math.Clamp(points, 0, 409);
        Active.Row(row).Height = points;
        State.RowHeights[row] = points;
        IsDirty = true;
    }

    public string GetEditText(int row, int column)
    {
        var cell = Active.Cell(row, column);
        if (cell.HasFormula)
            return EnsureEquals(cell.FormulaA1);

        return FormatRaw(cell.CachedValue);
    }

    public CellDisplay GetCellDisplay(int row, int column)
    {
        var cell = Active.Cell(row, column);
        if (cell.IsEmpty() && !cell.HasFormula)
            return new CellDisplay("", 11, false, false, false, CellHorizontalAlign.Left, null, null, null);

        var text = cell.HasFormula
            ? FormatRaw(cell.CachedValue)
            : FormatStored(cell);

        var font = cell.Style.Font;
        var alignment = cell.Style.Alignment;
        ReadColor(font.FontColor, out var red, out var green, out var blue);
        return new CellDisplay(
            text,
            font.FontSize,
            font.Bold,
            font.Italic,
            alignment.WrapText || text.Contains('\n'),
            HorizontalAlign(alignment.Horizontal, cell.CachedValue),
            red,
            green,
            blue);
    }

    public string GetValueKind(int row, int column)
    {
        var cell = Active.Cell(row, column);
        if (cell.HasFormula)
            return "formula";

        return cell.CachedValue.Type switch
        {
            XLDataType.Number => "number",
            XLDataType.Boolean => "bool",
            XLDataType.Text => "text",
            XLDataType.Error => "error",
            XLDataType.DateTime => "datetime",
            XLDataType.TimeSpan => "timespan",
            _ => "blank"
        };
    }

    public double GetNumber(int row, int column) => Active.Cell(row, column).CachedValue.GetNumber();

    public string GetFontName(int row, int column) => Active.Cell(row, column).Style.Font.FontName;

    public bool SetCellText(int row, int column, string text)
    {
        var current = GetEditText(row, column);
        if (string.Equals(current, text, StringComparison.Ordinal))
            return false;

        var cell = Active.Cell(row, column);
        var wasEmpty = cell.IsEmpty() && !cell.HasFormula;
        var trimmed = text.Trim();

        if (trimmed.Length == 0)
        {
            if (cell.IsEmpty() && !cell.HasFormula)
                return false;
            cell.Clear(XLClearOptions.Contents);
            IsDirty = true;
            return true;
        }

        if (trimmed.StartsWith('='))
        {
            cell.FormulaA1 = NormalizeFormula(trimmed);
        }
        else if (trimmed.StartsWith('\''))
        {
            cell.Value = trimmed[1..].Replace("\r\n", "\n").Replace('\r', '\n');
        }
        else if (TryParseNumber(trimmed, out var number))
        {
            cell.Value = number;
        }
        else if (trimmed.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
        {
            cell.Value = true;
        }
        else if (trimmed.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
        {
            cell.Value = false;
        }
        else
        {
            var stored = text.Replace("\r\n", "\n").Replace('\r', '\n');
            cell.Value = stored;
            if (stored.Contains('\n'))
                cell.Style.Alignment.WrapText = true;
        }

        cell.Style.Font.FontName = AppFonts.ExcelName;
        if (wasEmpty)
            cell.Style.Font.FontSize = 11;

        IsDirty = true;
        return true;
    }

    public void ClearCell(int row, int column)
    {
        var cell = Active.Cell(row, column);
        if (cell.IsEmpty() && !cell.HasFormula)
            return;

        cell.Clear(XLClearOptions.Contents);
        IsDirty = true;
    }

    public void Save()
    {
        if (string.IsNullOrEmpty(_path))
            throw new InvalidOperationException("Файл ещё не выбран.");

        _workbook.SaveAs(_path);
        IsDirty = false;
    }

    public void SaveAs(string path)
    {
        _workbook.SaveAs(path);
        _path = path;
        IsDirty = false;
    }

    public void Dispose() => _workbook.Dispose();

    public static string ColumnLetter(int column) => XLHelper.GetColumnLetterFromNumber(column);

    static SheetState LoadState(IXLWorksheet sheet)
    {
        var lastRow = sheet.LastRowUsed(XLCellsUsedOptions.Contents)?.RowNumber() ?? 0;
        var lastColumn = sheet.LastColumnUsed(XLCellsUsedOptions.Contents)?.ColumnNumber() ?? 0;
        var state = new SheetState
        {
            DisplayRows = Math.Max(lastRow + 20, MinimumRows),
            DisplayColumns = Math.Max(lastColumn + 5, MinimumColumns),
            DefaultColumnWidth = sheet.ColumnWidth,
            DefaultRowHeight = sheet.RowHeight
        };

        foreach (var (number, width) in ReadWidths(sheet))
            state.ColumnWidths[number] = width;
        foreach (var (number, height) in ReadHeights(sheet))
            state.RowHeights[number] = height;

        return state;
    }

    static IEnumerable<(int Number, double Width)> ReadWidths(IXLWorksheet sheet)
        => ReadMeasures(sheet, "ColumnsCollection", "Width");

    static IEnumerable<(int Number, double Height)> ReadHeights(IXLWorksheet sheet)
        => ReadMeasures(sheet, "RowsCollection", "Height");

    static IEnumerable<(int Number, double Measure)> ReadMeasures(IXLWorksheet sheet, string collectionName, string propertyName)
    {
        var internals = sheet.GetType().GetProperty("Internals", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
            ?.GetValue(sheet);
        if (internals?.GetType().GetProperty(collectionName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                ?.GetValue(internals) is not IEnumerable collection)
            yield break;

        foreach (var item in collection)
        {
            var pair = item.GetType();
            if (pair.GetProperty("Key")?.GetValue(item) is not int number)
                continue;
            var value = pair.GetProperty("Value")?.GetValue(item);
            if (value?.GetType().GetProperty(propertyName)?.GetValue(value) is double measure)
                yield return (number, measure);
        }
    }

    static string FormatStored(IXLCell cell)
    {
        try
        {
            return cell.GetFormattedString().Replace("\r\n", "\n");
        }
        catch
        {
            return FormatRaw(cell.CachedValue);
        }
    }

    static string FormatRaw(XLCellValue value)
    {
        return value.Type switch
        {
            XLDataType.Blank => "",
            XLDataType.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
            XLDataType.Number => value.GetNumber().ToString("G15", CultureInfo.CurrentCulture),
            XLDataType.Text => value.GetText(),
            XLDataType.DateTime => value.GetDateTime().ToString(CultureInfo.CurrentCulture),
            XLDataType.TimeSpan => value.GetTimeSpan().ToString(),
            XLDataType.Error => value.ToString(CultureInfo.CurrentCulture) ?? "",
            _ => value.ToString(CultureInfo.CurrentCulture) ?? ""
        };
    }

    static CellHorizontalAlign HorizontalAlign(XLAlignmentHorizontalValues horizontal, XLCellValue value)
    {
        return horizontal switch
        {
            XLAlignmentHorizontalValues.Center => CellHorizontalAlign.Center,
            XLAlignmentHorizontalValues.Right => CellHorizontalAlign.Right,
            XLAlignmentHorizontalValues.Left => CellHorizontalAlign.Left,
            _ => value.Type switch
            {
                XLDataType.Number or XLDataType.DateTime or XLDataType.TimeSpan => CellHorizontalAlign.Right,
                XLDataType.Boolean => CellHorizontalAlign.Center,
                _ => CellHorizontalAlign.Left
            }
        };
    }

    static void ReadColor(XLColor color, out byte? red, out byte? green, out byte? blue)
    {
        red = green = blue = null;
        try
        {
            if (color.ColorType != XLColorType.Color)
                return;
            var drawing = color.Color;
            red = drawing.R;
            green = drawing.G;
            blue = drawing.B;
        }
        catch
        {
            red = green = blue = null;
        }
    }

    static string EnsureEquals(string formula)
    {
        if (string.IsNullOrEmpty(formula))
            return "";
        return formula.StartsWith('=') ? formula : "=" + formula;
    }

    public static bool TryParseNumber(string text, out double number)
    {
        text = text.Trim();
        number = 0;
        if (text.Length == 0)
            return false;

        var culture = CultureInfo.CurrentCulture;
        if (culture.NumberFormat.NumberDecimalSeparator == ","
            && text.Contains('.')
            && !text.Contains(','))
        {
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)
                && double.IsFinite(number))
                return true;
        }

        if (double.TryParse(text, NumberStyles.Float, culture, out number) && double.IsFinite(number))
            return true;

        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)
            && double.IsFinite(number);
    }

    public static string NormalizeFormula(string text)
    {
        var body = text.StartsWith('=') ? text[1..] : text;
        var russian = body.Contains(';') || body.Any(c => c is >= '\u0400' and <= '\u04FF');
        if (!russian)
            return body;

        var builder = new StringBuilder(body.Length);
        var token = new StringBuilder();
        var inString = false;

        void Flush()
        {
            if (token.Length == 0)
                return;
            var word = token.ToString();
            builder.Append(RussianFunctions.TryGetValue(word, out var english) ? english : word);
            token.Clear();
        }

        for (var i = 0; i < body.Length; i++)
        {
            var c = body[i];
            if (c == '"')
            {
                Flush();
                inString = !inString;
                builder.Append(c);
                continue;
            }

            if (!inString && char.IsLetter(c))
            {
                token.Append(c);
                continue;
            }

            if (!inString && token.Length > 0 && char.IsDigit(c))
            {
                token.Append(c);
                continue;
            }

            Flush();
            if (!inString
                && c == ','
                && i > 0
                && i + 1 < body.Length
                && char.IsDigit(body[i - 1])
                && char.IsDigit(body[i + 1]))
            {
                builder.Append('.');
                continue;
            }

            builder.Append(!inString && c == ';' ? ',' : c);
        }

        Flush();
        return builder.ToString();
    }

    sealed class SheetState
    {
        public int DisplayRows { get; set; }

        public int DisplayColumns { get; set; }

        public double DefaultColumnWidth { get; init; }

        public double DefaultRowHeight { get; init; }

        public Dictionary<int, double> ColumnWidths { get; } = new();

        public Dictionary<int, double> RowHeights { get; } = new();
    }
}
