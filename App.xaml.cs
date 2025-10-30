using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Threading.Tasks;
using Gryzak.Views;

namespace Gryzak
{
    public partial class App : Application
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        private SplashWindow? _splashWindow;
        private MainWindow? _mainWindow;

        protected override void OnStartup(StartupEventArgs e)
        {
            // Alokuj konsolę na Windows, aby widzieć komunikaty Console.WriteLine
            AllocConsole();
            Console.WriteLine("=== Gryzak - Menedżer Zamówień ===");

            // Pokaż splash screen
            _splashWindow = new SplashWindow();
            _splashWindow.Show();

            // Załaduj główne okno w tle
            _ = LoadMainWindowAsync();

            base.OnStartup(e);
        }

        private async Task LoadMainWindowAsync()
        {
            // Symuluj ładowanie
            await Task.Delay(2000);

            // Utwórz główne okno
            _mainWindow = new MainWindow();
            
            // Zamknij splash screen
            _splashWindow?.CloseSplash();
            
            // Pokaż główne okno
            _mainWindow.Show();
        }
    }
}

