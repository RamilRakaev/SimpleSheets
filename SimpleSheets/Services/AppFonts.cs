using System.Windows;
using System.Windows.Media;

namespace SimpleSheets.Services;

public static class AppFonts
{
    public const string ExcelName = "Roboto";

    public static FontFamily Family { get; private set; } = new("Segoe UI");

    public static Typeface Regular { get; private set; } = new("Segoe UI");

    public static Typeface Bold { get; private set; } =
        new(new FontFamily("Segoe UI"), FontStyles.Normal, FontWeights.Bold, FontStretches.Normal);

    public static Typeface Italic { get; private set; } =
        new(new FontFamily("Segoe UI"), FontStyles.Italic, FontWeights.Normal, FontStretches.Normal);

    public static Typeface BoldItalic { get; private set; } =
        new(new FontFamily("Segoe UI"), FontStyles.Italic, FontWeights.Bold, FontStretches.Normal);

    public static void Load()
    {
        var family = new FontFamily(new Uri("pack://application:,,,/"), "./Fonts/#Roboto");
        var regular = new Typeface(family, FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
        if (!regular.TryGetGlyphTypeface(out _))
            return;

        Family = family;
        Regular = regular;
        Bold = new Typeface(family, FontStyles.Normal, FontWeights.Bold, FontStretches.Normal);
        Italic = new Typeface(family, FontStyles.Italic, FontWeights.Normal, FontStretches.Normal);
        BoldItalic = new Typeface(family, FontStyles.Italic, FontWeights.Bold, FontStretches.Normal);
    }
}
