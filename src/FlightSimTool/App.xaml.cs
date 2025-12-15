using System;
using System.Windows;

namespace FlightSimTool
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            ThemeManager.Initialize();
        }
    }
}
