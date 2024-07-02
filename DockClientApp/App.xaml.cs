using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace DockClientApp
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnExit(ExitEventArgs e)
        {
            Process[] processes = Process.GetProcessesByName("WINWORD");

            foreach (Process process in processes)
            {
                if (string.IsNullOrEmpty(process.MainWindowTitle))
                {
                    process.Kill();
                }
            }

            base.OnExit(e);
        }
    }
}
