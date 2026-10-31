using System.Windows;
using SimpleSheets.Services;

namespace SimpleSheets;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        AppFonts.Load();
        base.OnStartup(e);
    }
}
