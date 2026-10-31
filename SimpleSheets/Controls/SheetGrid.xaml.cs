using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using SimpleSheets.Services;

namespace SimpleSheets.Controls;

public partial class SheetGrid : UserControl
{
    static readonly Brush HeaderBrush = Freeze(new SolidColorBrush(Color.FromRgb(0xF3, 0xF3, 0xF3)));
    static readonly Brush HeaderSelectedBrush = Freeze(new SolidColorBrush(Color.FromRgb(0xD9, 0xED, 0xDF)));
    static readonly Brush GridLineBrush = Freeze(new SolidColorBrush(Color.FromRgb(0xE2, 0xE2, 0xE2)));
    static readonly Brush SelectionFill = Freeze(new SolidColorBrush(Color.FromRgb(0xE8, 0xF5, 0xEE)));
    static readonly Brush HeaderTextBrush = Freeze(new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)));
    static readonly Pen GridPen = FreezePen(new Pen(GridLineBrush, 1));
    static readonly Pen SelectionPen = FreezePen(new Pen(new SolidColorBrush(Color.FromRgb(0x21, 0x73, 0x46)), 2));
    static readonly Pen HeaderPen = FreezePen(new Pen(Freeze(new SolidColorBrush(Color.FromRgb(0xD0, 0xD0, 0xD0))), 1));

    readonly CellCanvas _cells;
    readonly HeaderCanvas _columns;
    readonly HeaderCanvas _rows;
    readonly TextBox _editor;

    WorkbookSession? _session;
    double[] _colPrefix = [0];
    double[] _rowPrefix = [0];
    double _mdw = 7;
    double _pixelsPerDip = 1;
    int _selectedRow = 1;
    int _selectedColumn = 1;
    bool _editing;
    bool _replaceMode;
    bool _suppressEditor;
    bool _scrollLock;

    int _dragIndex;
    bool _dragColumn;
    bool _dragMoved;
    double _dragAnchor;
    double _dragStartDips;

    public SheetGrid()
    {
        InitializeComponent();
        Focusable = true;

        _cells = new CellCanvas(this);
        CellHost.Child = _cells;
        _columns = new HeaderCanvas(this, columns: true);
        ColHost.Child = _columns;
        _rows = new HeaderCanvas(this, columns: false);
        RowHost.Child = _rows;

        _editor = new TextBox
        {
            Visibility = Visibility.Collapsed,
            BorderBrush = new SolidColorBrush(Color.FromRgb(0x21, 0x73, 0x46)),
            BorderThickness = new Thickness(2),
            Padding = new Thickness(2, 0, 2, 0),
            VerticalContentAlignment = VerticalAlignment.Center,
            AcceptsReturn = false,
            AcceptsTab = false
        };
        _cells.Children.Add(_editor);
        _editor.TextChanged += (_, _) =>
        {
            if (_suppressEditor || !_editing)
                return;
            EditTextChanged?.Invoke(this, EventArgs.Empty);
        };

        HScroll.ValueChanged += (_, _) => OnScrollChanged();
        VScroll.ValueChanged += (_, _) => OnScrollChanged();
        CellHost.SizeChanged += (_, _) =>
        {
            UpdateScrollbars();
            InvalidateAll();
        };
        Loaded += (_, _) => EnsureMetrics();
    }

    public event EventHandler? SelectionChanged;

    public event EventHandler? EditTextChanged;

    public event EventHandler? ContentChanged;

    public bool IsEditing => _editing;

    public int SelectedRow => _selectedRow;

    public int SelectedColumn => _selectedColumn;

    public string Address => WorkbookSession.ColumnLetter(_selectedColumn) + _selectedRow;

    public string EditText
    {
        get
        {
            if (_session == null)
                return "";
            return _editing ? _editor.Text : _session.GetEditText(_selectedRow, _selectedColumn);
        }
    }

    public bool IsEditorFocused => _editor.IsKeyboardFocusWithin;

    public void Attach(WorkbookSession session)
    {
        _session = session;
        _selectedRow = 1;
        _selectedColumn = 1;
        CancelEdit();
        _scrollLock = true;
        HScroll.Value = 0;
        VScroll.Value = 0;
        _scrollLock = false;
        EnsureMetrics();
        InvalidateAll();
        SelectionChanged?.Invoke(this, EventArgs.Empty);
        Focus();
    }

    public void ReloadSheet()
    {
        if (_editing)
            CommitEdit();
        _selectedRow = 1;
        _selectedColumn = 1;
        _scrollLock = true;
        HScroll.Value = 0;
        VScroll.Value = 0;
        _scrollLock = false;
        EnsureMetrics();
        InvalidateAll();
        SelectionChanged?.Invoke(this, EventArgs.Empty);
        Focus();
    }

    public void FocusGrid() => Focus();

    public void SetEditTextFromFormulaBar(string text)
    {
        if (_session == null)
            return;

        if (!_editing)
        {
            BeginEdit(text, replace: true, focusEditor: false);
            return;
        }

        if (_editor.Text == text)
            return;

        _suppressEditor = true;
        _editor.Text = text;
        if (_editor.IsKeyboardFocusWithin)
            _editor.CaretIndex = Math.Min(_editor.CaretIndex, text.Length);
        _suppressEditor = false;
    }

    public void CommitEdit()
    {
        if (!_editing || _session == null)
            return;

        var text = _editor.Text;
        var row = _selectedRow;
        var column = _selectedColumn;
        _editing = false;
        _editor.Visibility = Visibility.Collapsed;
        if (_session.SetCellText(row, column, text))
        {
            EnsureMetrics();
            ContentChanged?.Invoke(this, EventArgs.Empty);
        }

        InvalidateAll();
        EditTextChanged?.Invoke(this, EventArgs.Empty);
    }

    public void CancelEdit()
    {
        if (!_editing)
            return;

        _editing = false;
        _editor.Visibility = Visibility.Collapsed;
        InvalidateAll();
        EditTextChanged?.Invoke(this, EventArgs.Empty);
    }

    public void Move(int rowDelta, int columnDelta)
    {
        if (_session == null)
            return;

        if (_editing)
            CommitEdit();

        var row = Math.Max(1, _selectedRow + rowDelta);
        var column = Math.Max(1, _selectedColumn + columnDelta);
        Select(row, column);
    }

    void Select(int row, int column)
    {
        if (_session == null)
            return;

        if (_session.EnsureCapacity(row, column))
            RebuildMetrics();

        _selectedRow = Math.Clamp(row, 1, _session.DisplayRows);
        _selectedColumn = Math.Clamp(column, 1, _session.DisplayColumns);
        ScrollCellIntoView();
        InvalidateAll();
        SelectionChanged?.Invoke(this, EventArgs.Empty);
    }

    void BeginEdit(string? initial, bool replace, bool focusEditor = true)
    {
        if (_session == null)
            return;

        var display = _session.GetCellDisplay(_selectedRow, _selectedColumn);
        _editing = true;
        _replaceMode = replace;
        _editor.FontFamily = AppFonts.Family;
        _editor.FontSize = Math.Max(8, display.FontSize * 96d / 72d);
        _editor.FontWeight = display.Bold ? FontWeights.Bold : FontWeights.Normal;
        _editor.FontStyle = display.Italic ? FontStyles.Italic : FontStyles.Normal;
        _suppressEditor = true;
        _editor.Text = initial ?? _session.GetEditText(_selectedRow, _selectedColumn);
        _editor.CaretIndex = _editor.Text.Length;
        _suppressEditor = false;
        _editor.Visibility = Visibility.Visible;
        PositionEditor();
        if (focusEditor)
            _editor.Focus();
        EditTextChanged?.Invoke(this, EventArgs.Empty);
    }

    void EnsureMetrics()
    {
        var dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        if (_mdw <= 0 || Math.Abs(dpi - _pixelsPerDip) > 0.01)
        {
            _pixelsPerDip = dpi;
            _mdw = MeasureMaxDigitWidth();
        }

        if (_session != null)
            RebuildMetrics();
    }

    double MeasureMaxDigitWidth()
    {
        var max = 0d;
        var size = 11d * 96d / 72d;
        foreach (var digit in "0123456789")
        {
            var text = new FormattedText(
                digit.ToString(),
                CultureInfo.InvariantCulture,
                FlowDirection.LeftToRight,
                AppFonts.Regular,
                size,
                Brushes.Black,
                _pixelsPerDip);
            max = Math.Max(max, text.WidthIncludingTrailingWhitespace);
        }

        return max > 0.1 ? max : 7;
    }

    void RebuildMetrics()
    {
        if (_session == null)
            return;

        var columns = _session.DisplayColumns;
        var rows = _session.DisplayRows;
        _colPrefix = new double[columns + 1];
        _rowPrefix = new double[rows + 1];
        for (var column = 1; column <= columns; column++)
        {
            var dips = ExcelMetrics.WidthToDips(_session.GetColumnWidth(column), _mdw);
            _colPrefix[column] = _colPrefix[column - 1] + Math.Max(2, dips);
        }

        for (var row = 1; row <= rows; row++)
        {
            var dips = ExcelMetrics.PointsToDips(_session.GetRowHeight(row));
            _rowPrefix[row] = _rowPrefix[row - 1] + Math.Max(2, dips);
        }

        UpdateScrollbars();
    }

    void UpdateScrollbars()
    {
        if (_colPrefix.Length == 0 || _rowPrefix.Length == 0)
            return;

        _scrollLock = true;
        var viewWidth = Math.Max(0, CellHost.ActualWidth);
        var viewHeight = Math.Max(0, CellHost.ActualHeight);
        var totalWidth = _colPrefix[^1];
        var totalHeight = _rowPrefix[^1];
        HScroll.ViewportSize = viewWidth;
        VScroll.ViewportSize = viewHeight;
        HScroll.Maximum = Math.Max(0, totalWidth - viewWidth);
        VScroll.Maximum = Math.Max(0, totalHeight - viewHeight);
        HScroll.LargeChange = Math.Max(16, viewWidth * 0.9);
        VScroll.LargeChange = Math.Max(16, viewHeight * 0.9);
        HScroll.SmallChange = 40;
        VScroll.SmallChange = 24;
        if (HScroll.Value > HScroll.Maximum)
            HScroll.Value = HScroll.Maximum;
        if (VScroll.Value > VScroll.Maximum)
            VScroll.Value = VScroll.Maximum;
        _scrollLock = false;
        PositionEditor();
    }

    void OnScrollChanged()
    {
        if (_scrollLock)
            return;
        PositionEditor();
        InvalidateAll();
    }

    void InvalidateAll()
    {
        _cells.InvalidateVisual();
        _columns.InvalidateVisual();
        _rows.InvalidateVisual();
    }

    void ScrollCellIntoView()
    {
        if (_selectedColumn >= _colPrefix.Length || _selectedRow >= _rowPrefix.Length)
            return;

        var left = _colPrefix[_selectedColumn - 1];
        var right = _colPrefix[_selectedColumn];
        var top = _rowPrefix[_selectedRow - 1];
        var bottom = _rowPrefix[_selectedRow];
        var viewWidth = CellHost.ActualWidth;
        var viewHeight = CellHost.ActualHeight;

        if (left < HScroll.Value)
            HScroll.Value = left;
        else if (right > HScroll.Value + viewWidth)
            HScroll.Value = Math.Max(0, right - viewWidth);

        if (top < VScroll.Value)
            VScroll.Value = top;
        else if (bottom > VScroll.Value + viewHeight)
            VScroll.Value = Math.Max(0, bottom - viewHeight);
    }

    void PositionEditor()
    {
        if (!_editing || _selectedColumn >= _colPrefix.Length || _selectedRow >= _rowPrefix.Length)
            return;

        var x = _colPrefix[_selectedColumn - 1] - HScroll.Value;
        var y = _rowPrefix[_selectedRow - 1] - VScroll.Value;
        var width = Math.Max(8, _colPrefix[_selectedColumn] - _colPrefix[_selectedColumn - 1]);
        var height = Math.Max(8, _rowPrefix[_selectedRow] - _rowPrefix[_selectedRow - 1]);
        Canvas.SetLeft(_editor, x);
        Canvas.SetTop(_editor, y);
        _editor.Width = width;
        _editor.Height = height;
    }

    internal void RenderCells(DrawingContext context, Size size)
    {
        context.DrawRectangle(Brushes.White, null, new Rect(size));
        if (_session == null || _colPrefix.Length < 2 || _rowPrefix.Length < 2)
            return;

        var offsetX = HScroll.Value;
        var offsetY = VScroll.Value;
        var firstColumn = IndexAt(_colPrefix, offsetX);
        var lastColumn = IndexAt(_colPrefix, offsetX + size.Width);
        var firstRow = IndexAt(_rowPrefix, offsetY);
        var lastRow = IndexAt(_rowPrefix, offsetY + size.Height);

        for (var row = firstRow; row <= lastRow; row++)
        {
            var y = _rowPrefix[row - 1] - offsetY;
            var height = _rowPrefix[row] - _rowPrefix[row - 1];
            for (var column = firstColumn; column <= lastColumn; column++)
            {
                var x = _colPrefix[column - 1] - offsetX;
                var width = _colPrefix[column] - _colPrefix[column - 1];
                var rect = new Rect(x, y, width, height);
                var selected = row == _selectedRow && column == _selectedColumn;
                if (selected && !_editing)
                    context.DrawRectangle(SelectionFill, null, rect);

                var display = _session.GetCellDisplay(row, column);
                if (display.Text.Length > 0 && !(selected && _editing))
                    DrawCellText(context, display, rect);

                context.DrawLine(GridPen, new Point(rect.Right, rect.Top), new Point(rect.Right, rect.Bottom));
                context.DrawLine(GridPen, new Point(rect.Left, rect.Bottom), new Point(rect.Right, rect.Bottom));
                if (selected && !_editing)
                    context.DrawRectangle(null, SelectionPen, new Rect(rect.X + 1, rect.Y + 1, Math.Max(0, rect.Width - 2), Math.Max(0, rect.Height - 2)));
            }
        }
    }

    void DrawCellText(DrawingContext context, CellDisplay display, Rect rect)
    {
        var formatted = CreateText(display, Math.Max(1, rect.Width - 8));
        var y = rect.Bottom - formatted.Height - 3;
        if (y < rect.Top + 1)
            y = rect.Top + 1;

        context.PushClip(new RectangleGeometry(rect));
        context.DrawText(formatted, new Point(rect.X + 4, y));
        context.Pop();
    }

    FormattedText CreateText(CellDisplay display, double? maxWidth)
    {
        var brush = display.Red is byte red && display.Green is byte green && display.Blue is byte blue
            ? new SolidColorBrush(Color.FromRgb(red, green, blue))
            : Brushes.Black;
        var formatted = new FormattedText(
            display.Text,
            CultureInfo.CurrentCulture,
            FlowDirection.LeftToRight,
            TypefaceFor(display),
            Math.Max(1, display.FontSize * 96d / 72d),
            brush,
            _pixelsPerDip)
        {
            Trimming = TextTrimming.None,
            TextAlignment = display.Align switch
            {
                CellHorizontalAlign.Right => TextAlignment.Right,
                CellHorizontalAlign.Center => TextAlignment.Center,
                _ => TextAlignment.Left
            }
        };
        if (maxWidth is double width)
            formatted.MaxTextWidth = width;
        return formatted;
    }

    static Typeface TypefaceFor(CellDisplay display)
    {
        if (display.Bold && display.Italic)
            return AppFonts.BoldItalic;
        if (display.Bold)
            return AppFonts.Bold;
        if (display.Italic)
            return AppFonts.Italic;
        return AppFonts.Regular;
    }

    internal void RenderHeader(DrawingContext context, bool columns, Size size)
    {
        context.DrawRectangle(HeaderBrush, null, new Rect(size));
        if (_session == null)
            return;

        if (columns)
        {
            var first = IndexAt(_colPrefix, HScroll.Value);
            var last = IndexAt(_colPrefix, HScroll.Value + size.Width);
            for (var column = first; column <= last; column++)
            {
                var x = _colPrefix[column - 1] - HScroll.Value;
                var width = _colPrefix[column] - _colPrefix[column - 1];
                DrawHeaderCell(context, new Rect(x, 0, width, size.Height), WorkbookSession.ColumnLetter(column), column == _selectedColumn);
            }
        }
        else
        {
            var first = IndexAt(_rowPrefix, VScroll.Value);
            var last = IndexAt(_rowPrefix, VScroll.Value + size.Height);
            for (var row = first; row <= last; row++)
            {
                var y = _rowPrefix[row - 1] - VScroll.Value;
                var height = _rowPrefix[row] - _rowPrefix[row - 1];
                DrawHeaderCell(context, new Rect(0, y, size.Width, height), row.ToString(CultureInfo.CurrentCulture), row == _selectedRow);
            }
        }
    }

    void DrawHeaderCell(DrawingContext context, Rect rect, string text, bool selected)
    {
        if (selected)
            context.DrawRectangle(HeaderSelectedBrush, null, rect);

        var formatted = new FormattedText(
            text,
            CultureInfo.CurrentCulture,
            FlowDirection.LeftToRight,
            AppFonts.Regular,
            12,
            HeaderTextBrush,
            _pixelsPerDip);
        context.DrawText(formatted, new Point(rect.X + (rect.Width - formatted.Width) / 2, rect.Y + (rect.Height - formatted.Height) / 2));
        context.DrawLine(HeaderPen, new Point(rect.Right, rect.Top), new Point(rect.Right, rect.Bottom));
        context.DrawLine(HeaderPen, new Point(rect.Left, rect.Bottom), new Point(rect.Right, rect.Bottom));
    }

    internal void HeaderMouseDown(bool columns, MouseButtonEventArgs e, IInputElement source)
    {
        if (_session == null)
            return;

        var content = HeaderContentPosition(columns, e, source);
        var boundary = HitBoundary(columns, content);
        if (e.ClickCount == 2 && boundary > 0)
        {
            _dragIndex = 0;
            if (columns)
                AutofitColumn(boundary);
            else
                AutofitRow(boundary);
            e.Handled = true;
            return;
        }

        if (boundary > 0)
        {
            _dragIndex = boundary;
            _dragColumn = columns;
            _dragMoved = false;
            _dragAnchor = content;
            _dragStartDips = columns
                ? _colPrefix[boundary] - _colPrefix[boundary - 1]
                : _rowPrefix[boundary] - _rowPrefix[boundary - 1];
            Mouse.Capture((IInputElement)source);
            e.Handled = true;
            return;
        }

        if (_editing)
            CommitEdit();

        var index = columns
            ? IndexAt(_colPrefix, content)
            : IndexAt(_rowPrefix, content);
        if (columns)
            Select(_selectedRow, index);
        else
            Select(index, _selectedColumn);
        Focus();
        e.Handled = true;
    }

    internal void HeaderMouseMove(bool columns, MouseEventArgs e, FrameworkElement source)
    {
        if (_session == null)
            return;

        var content = HeaderContentPosition(columns, e, source);
        if (_dragIndex > 0 && _dragColumn == columns)
        {
            var delta = content - _dragAnchor;
            if (!_dragMoved && Math.Abs(delta) < 3)
                return;

            _dragMoved = true;
            var dips = Math.Max(8, _dragStartDips + delta);
            if (columns)
                _session.SetColumnWidth(_dragIndex, ExcelMetrics.DipsToWidth(dips, _mdw));
            else
                _session.SetRowHeight(_dragIndex, ExcelMetrics.DipsToPoints(dips));
            RebuildMetrics();
            InvalidateAll();
            return;
        }

        var boundary = HitBoundary(columns, content);
        source.Cursor = boundary > 0
            ? columns ? Cursors.SizeWE : Cursors.SizeNS
            : Cursors.Arrow;
    }

    internal void HeaderMouseUp(bool columns, IInputElement source)
    {
        if (_dragIndex == 0 || _dragColumn != columns)
            return;

        Mouse.Capture(null);
        var changed = _dragMoved;
        _dragIndex = 0;
        if (changed)
            ContentChanged?.Invoke(this, EventArgs.Empty);
    }

    double HeaderContentPosition(bool columns, MouseEventArgs e, IInputElement source)
    {
        var position = e.GetPosition(source);
        return columns ? position.X + HScroll.Value : position.Y + VScroll.Value;
    }

    int HitBoundary(bool columns, double content)
    {
        var prefix = columns ? _colPrefix : _rowPrefix;
        if (prefix.Length < 2)
            return 0;

        var index = IndexAt(prefix, content);
        if (Math.Abs(content - prefix[index]) <= 4)
            return index;
        if (index > 1 && Math.Abs(content - prefix[index - 1]) <= 4)
            return index - 1;
        return 0;
    }

    void AutofitColumn(int column)
    {
        if (_session == null)
            return;

        var last = _session.LastContentRow;
        var max = 0d;
        var any = false;
        for (var row = 1; row <= last; row++)
        {
            var display = _session.GetCellDisplay(row, column);
            if (display.Text.Length == 0)
                continue;
            any = true;
            var text = CreateText(display, null);
            max = Math.Max(max, text.WidthIncludingTrailingWhitespace + 8);
        }

        _session.SetColumnWidth(column, any
            ? ExcelMetrics.DipsToWidth(Math.Max(8, max), _mdw)
            : _session.DefaultColumnWidth);
        RebuildMetrics();
        InvalidateAll();
        ContentChanged?.Invoke(this, EventArgs.Empty);
    }

    void AutofitRow(int row)
    {
        if (_session == null)
            return;

        var last = _session.LastContentColumn;
        var maxPoints = 0d;
        var any = false;
        for (var column = 1; column <= last && column < _colPrefix.Length; column++)
        {
            var display = _session.GetCellDisplay(row, column);
            if (display.Text.Length == 0)
                continue;
            any = true;
            var width = _colPrefix[column] - _colPrefix[column - 1];
            var text = CreateText(display, display.WrapText ? Math.Max(1, width - 8) : null);
            maxPoints = Math.Max(maxPoints, ExcelMetrics.DipsToPoints(text.Height + 6));
        }

        _session.SetRowHeight(row, any ? Math.Max(10, maxPoints) : _session.DefaultRowHeight);
        RebuildMetrics();
        InvalidateAll();
        ContentChanged?.Invoke(this, EventArgs.Empty);
    }

    void CellMouseDown(MouseButtonEventArgs e)
    {
        if (_session == null)
            return;

        var position = e.GetPosition(_cells);
        var column = IndexAt(_colPrefix, position.X + HScroll.Value);
        var row = IndexAt(_rowPrefix, position.Y + VScroll.Value);
        if (e.ClickCount == 2)
        {
            if (row != _selectedRow || column != _selectedColumn)
                Select(row, column);
            BeginEdit(null, replace: false);
            e.Handled = true;
            return;
        }

        if (_editing)
            CommitEdit();
        Select(row, column);
        Focus();
        e.Handled = true;
    }

    protected override void OnPreviewKeyDown(KeyEventArgs e)
    {
        base.OnPreviewKeyDown(e);
        if (_session == null || e.Handled)
            return;

        if (_editing)
        {
            if (e.Key == Key.Enter && Keyboard.Modifiers == ModifierKeys.Alt)
            {
                var caret = _editor.CaretIndex;
                _editor.Text = _editor.Text.Insert(caret, "\n");
                _editor.CaretIndex = caret + 1;
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Enter)
            {
                CommitEdit();
                Move(Keyboard.Modifiers == ModifierKeys.Shift ? -1 : 1, 0);
                Focus();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Escape)
            {
                CancelEdit();
                Focus();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Tab)
            {
                CommitEdit();
                Move(0, Keyboard.Modifiers == ModifierKeys.Shift ? -1 : 1);
                Focus();
                e.Handled = true;
                return;
            }

            if (_replaceMode && e.Key is Key.Left or Key.Right or Key.Up or Key.Down)
            {
                CommitEdit();
                Move(e.Key == Key.Down ? 1 : e.Key == Key.Up ? -1 : 0, e.Key == Key.Right ? 1 : e.Key == Key.Left ? -1 : 0);
                Focus();
                e.Handled = true;
            }

            return;
        }

        if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.C)
        {
            Clipboard.SetText(EditText);
            e.Handled = true;
            return;
        }

        if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.V)
        {
            if (Clipboard.ContainsText())
            {
                _session.SetCellText(_selectedRow, _selectedColumn, Clipboard.GetText());
                EnsureMetrics();
                InvalidateAll();
                SelectionChanged?.Invoke(this, EventArgs.Empty);
                ContentChanged?.Invoke(this, EventArgs.Empty);
            }

            e.Handled = true;
            return;
        }

        if (e.Key is Key.Delete or Key.Back)
        {
            _session.ClearCell(_selectedRow, _selectedColumn);
            InvalidateAll();
            SelectionChanged?.Invoke(this, EventArgs.Empty);
            ContentChanged?.Invoke(this, EventArgs.Empty);
            e.Handled = true;
            return;
        }

        if (e.Key == Key.F2)
        {
            BeginEdit(null, replace: false);
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Enter)
        {
            Move(Keyboard.Modifiers == ModifierKeys.Shift ? -1 : 1, 0);
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Tab)
        {
            Move(0, Keyboard.Modifiers == ModifierKeys.Shift ? -1 : 1);
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Home && Keyboard.Modifiers == ModifierKeys.Control)
        {
            Select(1, 1);
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Home)
        {
            Select(_selectedRow, 1);
            e.Handled = true;
            return;
        }

        var page = Math.Max(1, (int)(CellHost.ActualHeight / 22) - 1);
        switch (e.Key)
        {
            case Key.Left:
                Move(0, -1);
                e.Handled = true;
                break;
            case Key.Right:
                Move(0, 1);
                e.Handled = true;
                break;
            case Key.Up:
                Move(-1, 0);
                e.Handled = true;
                break;
            case Key.Down:
                Move(1, 0);
                e.Handled = true;
                break;
            case Key.PageDown:
                Move(page, 0);
                e.Handled = true;
                break;
            case Key.PageUp:
                Move(-page, 0);
                e.Handled = true;
                break;
        }
    }

    protected override void OnPreviewTextInput(TextCompositionEventArgs e)
    {
        base.OnPreviewTextInput(e);
        if (_editing || _session == null || e.Handled)
            return;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control) || Keyboard.Modifiers.HasFlag(ModifierKeys.Alt))
            return;
        if (e.Text.Length == 0 || char.IsControl(e.Text, 0))
            return;

        BeginEdit(e.Text, replace: true);
        e.Handled = true;
    }

    protected override void OnPreviewMouseWheel(MouseWheelEventArgs e)
    {
        base.OnPreviewMouseWheel(e);
        if (_session == null)
            return;

        var bar = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) ? HScroll : VScroll;
        var next = bar.Value - e.Delta / 120d * 48;
        bar.Value = Math.Clamp(next, bar.Minimum, bar.Maximum);
        e.Handled = true;
    }

    static int IndexAt(double[] prefix, double position)
    {
        if (prefix.Length <= 1)
            return 1;
        if (position <= 0)
            return 1;

        var lo = 1;
        var hi = prefix.Length - 1;
        while (lo <= hi)
        {
            var mid = (lo + hi) / 2;
            if (position < prefix[mid - 1])
                hi = mid - 1;
            else if (position >= prefix[mid])
                lo = mid + 1;
            else
                return mid;
        }

        return Math.Clamp(hi, 1, prefix.Length - 1);
    }

    static T Freeze<T>(T freezable) where T : Freezable
    {
        freezable.Freeze();
        return freezable;
    }

    static Pen FreezePen(Pen pen)
    {
        pen.Freeze();
        return pen;
    }

    sealed class CellCanvas : Canvas
    {
        readonly SheetGrid _owner;

        public CellCanvas(SheetGrid owner)
        {
            _owner = owner;
            Background = Brushes.White;
            ClipToBounds = true;
        }

        protected override void OnRender(DrawingContext drawingContext) => _owner.RenderCells(drawingContext, RenderSize);

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e) => _owner.CellMouseDown(e);
    }

    sealed class HeaderCanvas : Canvas
    {
        readonly SheetGrid _owner;
        readonly bool _columns;

        public HeaderCanvas(SheetGrid owner, bool columns)
        {
            _owner = owner;
            _columns = columns;
            Background = HeaderBrush;
            ClipToBounds = true;
        }

        protected override void OnRender(DrawingContext drawingContext) => _owner.RenderHeader(drawingContext, _columns, RenderSize);

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e) => _owner.HeaderMouseDown(_columns, e, this);

        protected override void OnMouseMove(MouseEventArgs e) => _owner.HeaderMouseMove(_columns, e, this);

        protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e) => _owner.HeaderMouseUp(_columns, this);

        protected override void OnMouseLeave(MouseEventArgs e)
        {
            if (_owner._dragIndex == 0)
                Cursor = Cursors.Arrow;
        }
    }
}
