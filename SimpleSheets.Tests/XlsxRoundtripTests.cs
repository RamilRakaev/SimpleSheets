using System.Globalization;
using ClosedXML.Excel;
using SimpleSheets.Services;
using Xunit;

namespace SimpleSheets.Tests;

public class XlsxRoundtripTests
{
    [Fact]
    public void Save_keeps_values_formula_sizes_font_and_other_sheets()
    {
        var previous = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ru-RU");
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            using (var seed = new XLWorkbook())
            {
                var sheet = seed.AddWorksheet("Лист1");
                sheet.Cell(1, 1).Value = "Keep";
                sheet.Cell(1, 1).Style.Font.FontName = "Calibri";
                sheet.Column(1).Width = 18;
                sheet.Row(1).Height = 24;
                seed.AddWorksheet("Второй").Cell(1, 1).Value = "stay";
                seed.SaveAs(path);
            }

            using (var session = WorkbookSession.Open(path))
            {
                Assert.Equal("Calibri", session.GetFontName(1, 1));
                Assert.Equal(18, session.GetColumnWidth(1), 2);
                Assert.Equal(24, session.GetRowHeight(1), 2);

                Assert.True(session.SetCellText(2, 1, "Новое"));
                Assert.True(session.SetCellText(2, 2, "12,5"));
                Assert.True(session.SetCellText(3, 1, "=СУММ(A1;B2)"));
                session.SetColumnWidth(2, 15);
                session.SetRowHeight(3, 40);
                session.Save();
            }

            using var saved = new XLWorkbook(path);
            var first = saved.Worksheet("Лист1");
            Assert.Equal("Keep", first.Cell(1, 1).GetString());
            Assert.Equal("Calibri", first.Cell(1, 1).Style.Font.FontName);
            Assert.Equal("Новое", first.Cell(2, 1).GetString());
            Assert.Equal("Roboto", first.Cell(2, 1).Style.Font.FontName);
            Assert.Equal(XLDataType.Number, first.Cell(2, 2).DataType);
            Assert.Equal(12.5, first.Cell(2, 2).GetDouble(), 5);
            Assert.Equal("Roboto", first.Cell(2, 2).Style.Font.FontName);
            Assert.Equal("SUM(A1,B2)", first.Cell(3, 1).FormulaA1);
            Assert.Equal("Roboto", first.Cell(3, 1).Style.Font.FontName);
            Assert.Equal(18, first.Column(1).Width, 2);
            Assert.Equal(15, first.Column(2).Width, 2);
            Assert.Equal(24, first.Row(1).Height, 2);
            Assert.Equal(40, first.Row(3).Height, 2);
            Assert.Equal("stay", saved.Worksheet("Второй").Cell(1, 1).GetString());
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void New_workbook_starts_clean_with_default_sizes()
    {
        using var session = WorkbookSession.CreateNew();
        Assert.False(session.IsDirty);
        Assert.Equal("", session.GetEditText(1, 1));
        Assert.InRange(session.GetColumnWidth(1), 8, 9);
        Assert.InRange(session.GetRowHeight(1), 14, 16);
        Assert.Equal(session.GetColumnWidth(1), session.GetColumnWidth(3));
        Assert.Equal(session.GetRowHeight(1), session.GetRowHeight(2));
    }

    [Fact]
    public void Width_conversion_roundtrips_near_the_original_excel_width()
    {
        const double mdw = 7;
        var dips = ExcelMetrics.WidthToDips(12.5, mdw);
        var width = ExcelMetrics.DipsToWidth(dips, mdw);
        Assert.InRange(Math.Abs(width - 12.5), 0, 0.1);
    }

    [Fact]
    public void NormalizeFormula_keeps_english_argument_commas()
    {
        Assert.Equal("SUM(1,2)", WorkbookSession.NormalizeFormula("=SUM(1,2)"));
        Assert.Equal("SUM(1.5,2)", WorkbookSession.NormalizeFormula("=СУММ(1,5;2)"));
        Assert.Equal("\"a;b\"", WorkbookSession.NormalizeFormula("=\"a;b\""));
    }
}
