using System;
using System.Windows;
using System.Threading.Tasks;
using Gryzak.Views;
using static Gryzak.Services.Logger;

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
            
            Info("=== Gryzak - Menedżer Zamówień ===", "App");

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
            Critical(e.Exception, "App", "Nieobsłużony wyjątek w wątku UI");
            ZwolnijLicencjeSubiekta();
            
            // Można pozwolić aplikacji kontynuować lub zakończyć działanie
            // e.Handled = true; // Odkomentuj jeśli chcesz kontynuować mimo błędu
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Critical($"Nieobsłużony wyjątek w domenie aplikacji: {e.ExceptionObject}", "App");
            ZwolnijLicencjeSubiekta();
        }

        private void ZwolnijLicencjeSubiekta()
        {
            try
            {
                Info("Zwalnianie licencji Subiekta GT z powodu błędu...", "App");
                var subiektService = new Gryzak.Services.SubiektService();
                subiektService.ZwolnijLicencje();
            }
            catch (Exception ex)
            {
                Error(ex, "App", "Błąd podczas zwalniania licencji przy błędzie aplikacji");
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
                Info("Zamykanie okna debugowania...", "App");
                DebugWindow.CloseWindowAndThread();
            }
            catch (Exception ex)
            {
                Error(ex, "App", "Błąd podczas zamykania okna debugowania");
            }

            base.OnExit(e);
        }

        private async Task LoadMainWindowAsync()
        {
            try
            {
                // Utwórz główne okno (ale jeszcze nie pokazuj)
                _mainWindow = new MainWindow();
                
                // Pobierz MainViewModel z MainWindow
                if (_mainWindow.DataContext is ViewModels.MainViewModel mainViewModel && _splashWindow != null)
                {
                    // Aktualizuj splash screen podczas ładowania
                    _splashWindow.UpdateProgress(10, "Inicjalizacja aplikacji...");
                    
                    // Przekaż splash screen do MainViewModel aby mógł aktualizować postęp
                    mainViewModel.SetSplashWindow(_splashWindow);
                    
                    // Załaduj zamówienia (to zajmie trochę czasu)
                    // MainViewModel będzie aktualizował splash screen wewnątrz LoadOrdersAsync
                    _splashWindow.UpdateProgress(20, "Ładowanie zamówień z API...");
                    await mainViewModel.LoadOrdersAsync(false);
                    
                    // Sprawdź czy zamówienia zostały załadowane
                    _splashWindow.UpdateProgress(90, "Finalizacja...");
                    await Task.Delay(300); // Krótkie opóźnienie dla płynności
                    
                    _splashWindow.UpdateProgress(100, "Gotowe!");
                    await Task.Delay(300);
                }
                else
                {
                    // Fallback - jeśli nie ma MainViewModel, poczekaj chwilę
                    _splashWindow?.UpdateProgress(50, "Przygotowywanie interfejsu...");
                    await Task.Delay(1000);
                    _splashWindow?.UpdateProgress(100, "Gotowe!");
                    await Task.Delay(300);
                }
            }
            catch (Exception ex)
            {
                Error(ex, "App", "Błąd podczas ładowania głównego okna");
                _splashWindow?.UpdateProgress(0, "Błąd ładowania...");
                await Task.Delay(1000);
            }
            finally
            {
                // Zamknij splash screen
                _splashWindow?.CloseSplash();
                
                // Pokaż główne okno (które ma już załadowane zamówienia)
                if (_mainWindow != null)
                {
                    // Tymczasowo ustaw Topmost, aby upewnić się że okno będzie na pierwszym planie
                    _mainWindow.Topmost = true;
                    _mainWindow.Show();
                    _mainWindow.Activate();
                    _mainWindow.Focus();
                    _mainWindow.WindowState = WindowState.Normal;
                    _mainWindow.BringIntoView();
                    
                    // Po aktywacji wyłącz Topmost (żeby nie było zawsze na wierzchu)
                    // Użyj Task.Delay aby zrobić to asynchronicznie po zakończeniu renderowania
                    _ = Task.Run(async () =>
                    {
                        await Task.Delay(100); // Krótkie opóźnienie aby okno się w pełni wyświetliło
                        _mainWindow.Dispatcher.Invoke(() =>
                        {
                            _mainWindow.Topmost = false;
                        });
                    });
                }
            }
        }
    }
}

