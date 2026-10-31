namespace SimpleSheets.Services;

public static class ExcelMetrics
{
    public static double WidthToDips(double width, double maxDigitWidth)
        => Math.Truncate(width * maxDigitWidth + 5);

    public static double DipsToWidth(double dips, double maxDigitWidth)
    {
        if (maxDigitWidth <= 0)
            return 8.43;

        var width = Math.Truncate((dips - 5) / maxDigitWidth * 256d) / 256d;
        return width < 0 ? 0 : width;
    }

    public static double PointsToDips(double points) => points * 96d / 72d;

    public static double DipsToPoints(double dips) => dips * 72d / 96d;
}
