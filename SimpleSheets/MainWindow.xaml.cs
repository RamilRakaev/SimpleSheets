using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using SimpleSheets.Services;

namespace SimpleSheets;

public partial class MainWindow : Window
{
    WorkbookSession? _session;
    bool _sync;
    bool _sheetSync;

    public MainWindow()
    {
        InitializeComponent();
        FontFamily = AppFonts.Family;
        FormulaBox.FontFamily = AppFonts.Family;
        AddressBox.FontFamily = AppFonts.Family;
        Sheet.FontFamily = AppFonts.Family;
        Sheet.SelectionChanged += (_, _) => ShowSelection();
        Sheet.EditTextChanged += (_, _) => ShowSelection();
        Sheet.ContentChanged += (_, _) => UpdateTitle();
        Loaded += (_, _) =>
        {
            NewBook();
            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Input, new Action(Sheet.FocusGrid));
        };
        Closing += OnClosing;
    }

    void NewBook()
    {
        ReplaceSession(WorkbookSession.CreateNew());
    }

    void ReplaceSession(WorkbookSession session)
    {
        _session?.Dispose();
        _session = session;
        ReloadSheets();
        Sheet.Attach(session);
        UpdateTitle();
        ShowSelection();
    }

    void ReloadSheets()
    {
        if (_session == null)
            return;

        _sheetSync = true;
        SheetTabs.Items.Clear();
        foreach (var name in _session.SheetNames)
            SheetTabs.Items.Add(name);
        SheetTabs.SelectedIndex = _session.ActiveSheetIndex;
        _sheetSync = false;
    }

    void ShowSelection()
    {
        if (_sync)
            return;

        _sync = true;
        AddressBox.Text = Sheet.Address;
        var text = Sheet.EditText;
        if (FormulaBox.Text != text)
            FormulaBox.Text = text;
        _sync = false;
    }

    void UpdateTitle()
    {
        if (_session == null)
            return;
        var name = _session.Title;
        Title = _session.IsDirty ? $"*{name} — SimpleSheets" : $"{name} — SimpleSheets";
    }

    void NewClick(object sender, RoutedEventArgs e)
    {
        if (!ConfirmDiscard())
            return;
        NewBook();
    }

    void OpenClick(object sender, RoutedEventArgs e)
    {
        if (!ConfirmDiscard())
            return;

        var dialog = new OpenFileDialog
        {
            Filter = "Книги Excel (*.xlsx)|*.xlsx",
            Title = "Открыть"
        };
        if (dialog.ShowDialog(this) != true)
            return;

        try
        {
            ReplaceSession(WorkbookSession.Open(dialog.FileName));
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Не удалось открыть файл.\n" + ex.Message, "SimpleSheets", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    void SaveClick(object sender, RoutedEventArgs e) => Save();

    void SaveAsClick(object sender, RoutedEventArgs e) => SaveAs();

    void ExitClick(object sender, RoutedEventArgs e) => Close();

    bool Save()
    {
        Sheet.CommitEdit();
        if (_session == null)
            return false;
        if (string.IsNullOrEmpty(_session.FilePath))
            return SaveAs();

        try
        {
            _session.Save();
            UpdateTitle();
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Не удалось сохранить файл.\n" + ex.Message, "SimpleSheets", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
    }

    bool SaveAs()
    {
        Sheet.CommitEdit();
        if (_session == null)
            return false;

        var dialog = new SaveFileDialog
        {
            Filter = "Книги Excel (*.xlsx)|*.xlsx",
            Title = "Сохранить как",
            FileName = _session.FilePath == null ? "Книга1.xlsx" : System.IO.Path.GetFileName(_session.FilePath),
            DefaultExt = ".xlsx",
            AddExtension = true
        };
        if (dialog.ShowDialog(this) != true)
            return false;

        try
        {
            _session.SaveAs(dialog.FileName);
            UpdateTitle();
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Не удалось сохранить файл.\n" + ex.Message, "SimpleSheets", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
    }

    bool ConfirmDiscard()
    {
        Sheet.CommitEdit();
        if (_session is not { IsDirty: true })
            return true;

        var result = MessageBox.Show(
            this,
            "Сохранить изменения?",
            "SimpleSheets",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question);
        return result switch
        {
            MessageBoxResult.Cancel => false,
            MessageBoxResult.Yes => Save(),
            _ => true
        };
    }

    void FormulaBoxTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_sync)
            return;
        _sync = true;
        Sheet.SetEditTextFromFormulaBar(FormulaBox.Text);
        _sync = false;
    }

    void FormulaBoxPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            Sheet.CommitEdit();
            Sheet.Move(Keyboard.Modifiers == ModifierKeys.Shift ? -1 : 1, 0);
            Sheet.FocusGrid();
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            Sheet.CancelEdit();
            ShowSelection();
            Sheet.FocusGrid();
            e.Handled = true;
        }
    }

    void FormulaBoxLostFocus(object sender, RoutedEventArgs e)
    {
        if (Sheet.IsEditing && !Sheet.IsEditorFocused)
            Sheet.CommitEdit();
    }

    void SheetTabsSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_sheetSync || _session == null || SheetTabs.SelectedIndex < 0)
            return;
        if (SheetTabs.SelectedIndex == _session.ActiveSheetIndex)
            return;

        Sheet.CommitEdit();
        _session.SetActiveSheet(SheetTabs.SelectedIndex);
        Sheet.ReloadSheet();
        UpdateTitle();
    }

    void OnClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (!ConfirmDiscard())
            e.Cancel = true;
        else
            _session?.Dispose();
    }

    protected override void OnPreviewKeyDown(KeyEventArgs e)
    {
        base.OnPreviewKeyDown(e);
        if (e.Handled)
            return;

        if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.S)
        {
            Save();
            e.Handled = true;
        }
        else if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.O)
        {
            OpenClick(this, e);
            e.Handled = true;
        }
        else if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.N)
        {
            NewClick(this, e);
            e.Handled = true;
        }
    }
}
