using System;
using System.Windows;
using System.Threading.Tasks;
using Gryzak.Views;

namespace Gryzak
{
    public partial class App : Application
    {

        private SplashWindow? _splashWindow;
        private MainWindow? _mainWindow;
        private DebugWindow? _debugWindow;

        protected override void OnStartup(StartupEventArgs e)
        {
            // Utwórz okno debugowania (niewidoczne) i zacznij przechwytywać logi od razu
            // Nie używamy już AllocConsole() - wszystkie logi idą do okna debugowania WPF
            // Okno jest tworzone w osobnym wątku, więc jest całkowicie niezależne
            _debugWindow = DebugWindow.GetOrCreateInstance();
            _debugWindow.HideWindow(); // Ukryj okno na starcie
            _debugWindow.StartCapturing();
            
            Console.WriteLine("=== Gryzak - Menedżer Zamówień ===");

            // Zarejestruj handlery dla nieobsłużonych wyjątków, aby zwolnić licencję
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;

            // Pokaż splash screen
            _splashWindow = new SplashWindow();
            _splashWindow.Show();

            // Załaduj główne okno w tle
            _ = LoadMainWindowAsync();

            base.OnStartup(e);
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            Console.WriteLine($"[App] NIEOBSŁUŻONY WYJĄTEK w wątku UI: {e.Exception}");
            ZwolnijLicencjeSubiekta();
            
            // Można pozwolić aplikacji kontynuować lub zakończyć działanie
            // e.Handled = true; // Odkomentuj jeśli chcesz kontynuować mimo błędu
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Console.WriteLine($"[App] NIEOBSŁUŻONY WYJĄTEK w domenie aplikacji: {e.ExceptionObject}");
            ZwolnijLicencjeSubiekta();
        }

        private void ZwolnijLicencjeSubiekta()
        {
            try
            {
                Console.WriteLine("[App] Zwalnianie licencji Subiekta GT z powodu błędu...");
                var subiektService = new Gryzak.Services.SubiektService();
                subiektService.ZwolnijLicencje();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[App] Błąd podczas zwalniania licencji przy błędzie aplikacji: {ex.Message}");
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // Zwolnij licencję Subiekta GT przed zamknięciem aplikacji
            // (backup na wypadek zamknięcia w inny sposób niż przez MainWindow)
            ZwolnijLicencjeSubiekta();
            
            // Zamknij okno debugowania i jego wątek
            try
            {
                Console.WriteLine("[App] Zamykanie okna debugowania...");
                DebugWindow.CloseWindowAndThread();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[App] Błąd podczas zamykania okna debugowania: {ex.Message}");
            }

            base.OnExit(e);
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

