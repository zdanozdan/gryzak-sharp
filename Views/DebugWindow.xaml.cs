using System;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Threading;

namespace Gryzak.Views
{
    public partial class DebugWindow : Window
    {
        private DebugTextWriter? _debugWriter;
        private static DebugWindow? _instance;
        private static Thread? _windowThread;
        private static AutoResetEvent? _windowReady;
        private StreamWriter? _logFileWriter;
        private string? _logFilePath;

        public DebugWindow()
        {
            InitializeComponent();
            _instance = this;
            
            // Upewnij się, że okno jest całkowicie niezależne
            this.Owner = null;
            this.ShowInTaskbar = true;
        }

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            UpdateStatus();
        }

        private void UpdateStatus()
        {
            if (LogTextBox != null)
            {
                var lines = LogTextBox.Text.Split('\n');
                int lineCount = lines.Length;
                StatusTextBlock.Text = $"Liczba linii: {lineCount} | Rozmiar: {LogTextBox.Text.Length} znaków";
            }
        }

        public void AppendLog(string message)
        {
            try
            {
                // Zawsze używaj BeginInvoke, nawet jeśli jesteśmy w tym samym wątku
                // aby zapewnić bezpieczeństwo wielowątkowości
                if (LogTextBox != null)
                {
                    if (LogTextBox.Dispatcher.CheckAccess())
                    {
                        LogTextBox.AppendText(message);
                        LogTextBox.ScrollToEnd();
                        UpdateStatus();
                    }
                    else
                    {
                        LogTextBox.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            LogTextBox?.AppendText(message);
                            LogTextBox?.ScrollToEnd();
                            UpdateStatus();
                        }));
                    }
                }
            }
            catch (Exception ex)
            {
                // Ignoruj błędy - może okno nie jest jeszcze gotowe
                System.Diagnostics.Debug.WriteLine($"Błąd AppendLog: {ex.Message}");
            }
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            LogTextBox.Clear();
            UpdateStatus();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            // Jeśli aplikacja się zamyka, pozwól na zamknięcie okna
            // W przeciwnym razie tylko ukryj okno
            if (_isShuttingDown)
            {
                // Aplikacja się zamyka - pozwól zamknąć okno
                e.Cancel = false;
            }
            else
            {
                // Aplikacja nie zamyka się - tylko ukryj okno
                e.Cancel = true;
                this.Hide();
            }
        }
        
        public void ShowWindow()
        {
            if (Dispatcher.CheckAccess())
            {
                this.Show();
                this.Activate();
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    this.Show();
                    this.Activate();
                });
            }
        }
        
        public void HideWindow()
        {
            if (Dispatcher.CheckAccess())
            {
                this.Hide();
            }
            else
            {
                Dispatcher.Invoke(() => this.Hide());
            }
        }

        public static DebugWindow GetOrCreateInstance()
        {
            if (_instance == null)
            {
                // Jeśli okno nie istnieje, utwórz je w osobnym wątku
                // Dzięki temu będzie niezależne od modalnych okien
                CreateWindowInSeparateThread();
                
                // Poczekaj aż okno zostanie utworzone
                _windowReady?.WaitOne(5000);
            }
            return _instance!;
        }
        
        private static void CreateWindowInSeparateThread()
        {
            _windowReady = new AutoResetEvent(false);
            
            _windowThread = new Thread(() =>
            {
                // Utwórz nowy dispatcher dla tego wątku
                _instance = new DebugWindow();
                _instance.Show();
                
                // Poinformuj główny wątek że okno jest gotowe
                _windowReady.Set();
                
                // Uruchom dispatcher loop dla tego wątku
                Dispatcher.Run();
            });
            
            _windowThread.SetApartmentState(ApartmentState.STA);
            _windowThread.IsBackground = true; // Wątek powinien kończyć się z aplikacją
            _windowThread.Start();
        }
        
        private static bool _isShuttingDown = false;
        
        public static void CloseWindowAndThread()
        {
            _isShuttingDown = true;
            
            if (_instance != null)
            {
                // Zatrzymaj przechwytywanie logów
                _instance.StopCapturing();
                
                // Zamknij okno przez dispatcher
                if (_instance.Dispatcher != null && !_instance.Dispatcher.HasShutdownStarted)
                {
                    try
                    {
                        _instance.Dispatcher.Invoke(() =>
                        {
                            try
                            {
                                _instance.Close();
                            }
                            catch { }
                        });
                    }
                    catch { }
                }
                
                // Zatrzymaj logowanie do pliku przed zamknięciem
                _instance.StopLoggingToFile();
                
                // Teraz zamknij dispatcher w wątku okna
                if (_instance.Dispatcher != null && !_instance.Dispatcher.HasShutdownStarted)
                {
                    try
                    {
                        _instance.Dispatcher.BeginInvokeShutdown(System.Windows.Threading.DispatcherPriority.Normal);
                    }
                    catch { }
                }
                
                _instance = null;
            }
            
            // Poczekaj aż wątek się zakończy (max 2 sekundy)
            if (_windowThread != null && _windowThread.IsAlive)
            {
                _windowThread.Join(2000);
            }
            
            _windowThread = null;
        }

        // Klasa pomocnicza do przechwytywania Console.WriteLine
        private class DebugTextWriter : TextWriter
        {
            private readonly DebugWindow _window;
            private readonly TextWriter _originalOut;

            public DebugTextWriter(DebugWindow window)
            {
                _window = window;
                _originalOut = Console.Out;
            }

            public override Encoding Encoding => Encoding.UTF8;

            public override void Write(char value)
            {
                _originalOut.Write(value);
                _window.AppendLog(value.ToString());
                _window.WriteToLogFile(value.ToString());
            }

            public override void Write(string? value)
            {
                _originalOut.Write(value);
                if (value != null)
                {
                    _window.AppendLog(value);
                    _window.WriteToLogFile(value);
                }
            }

            public override void WriteLine(string? value)
            {
                _originalOut.WriteLine(value);
                if (value != null)
                {
                    string lineWithNewline = value + Environment.NewLine;
                    _window.AppendLog(lineWithNewline);
                    _window.WriteToLogFile(lineWithNewline);
                }
                else
                {
                    _window.AppendLog(Environment.NewLine);
                    _window.WriteToLogFile(Environment.NewLine);
                }
            }

            public TextWriter OriginalOut => _originalOut;
        }

        public void StartCapturing()
        {
            if (_debugWriter == null)
            {
                _debugWriter = new DebugTextWriter(this);
                Console.SetOut(_debugWriter);
            }
        }

        public void StopCapturing()
        {
            if (_debugWriter != null)
            {
                // Przywróć oryginalny output (konsola)
                Console.SetOut(_debugWriter.OriginalOut);
                _debugWriter = null;
            }
            
            // Zamknij plik logów jeśli jest otwarty
            StopLoggingToFile();
        }
        
        private void LogToFileCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            StartLoggingToFile();
        }
        
        private void LogToFileCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            StopLoggingToFile();
        }
        
        private void StartLoggingToFile()
        {
            try
            {
                // Zawsze zapisuj logi w katalogu "logs" w katalogu projektu (obok bin)
                // W trybie deweloperskim: [projekt]\logs
                // W produkcji: [instalacja]\logs
                string logDir;
                
                try
                {
                    // W trybie deweloperskim: znajdź katalog projektu przez przejście w górę z katalogu bin
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    string? assemblyLocation = Path.GetDirectoryName(assembly.Location);
                    
                    if (assemblyLocation != null && (assemblyLocation.Contains("\\bin\\") || assemblyLocation.Contains("/bin/")))
                    {
                        // Jesteśmy w katalogu bin - przejdź do katalogu projektu
                        var binIndex = assemblyLocation.LastIndexOf("\\bin\\");
                        if (binIndex == -1)
                            binIndex = assemblyLocation.LastIndexOf("/bin/");
                        
                        if (binIndex >= 0)
                        {
                            // Przejdź do katalogu projektu (rodzica bin)
                            string projectDir = assemblyLocation.Substring(0, binIndex);
                            logDir = Path.Combine(projectDir, "logs");
                        }
                        else
                        {
                            // Nie znaleziono bin - użyj katalogu wykonywalnego
                            logDir = Path.Combine(assemblyLocation, "logs");
                        }
                    }
                    else
                    {
                        // Produkcja lub inna konfiguracja - użyj katalogu wykonywalnego
                        logDir = Path.Combine(assemblyLocation ?? AppDomain.CurrentDomain.BaseDirectory, "logs");
                    }
                }
                catch
                {
                    // Fallback: użyj katalogu roboczego lub BaseDirectory
                    string baseDir = Environment.CurrentDirectory;
                    logDir = Path.Combine(baseDir, "logs");
                }
                
                Directory.CreateDirectory(logDir);
                
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                _logFilePath = Path.Combine(logDir, $"gryzak_log_{timestamp}.txt");
                
                // Otwórz plik do zapisu (append mode)
                _logFileWriter = new StreamWriter(_logFilePath, append: true, Encoding.UTF8)
                {
                    AutoFlush = true
                };
                
                // Zapisz nagłówek
                _logFileWriter.WriteLine($"=== Gryzak Debug Log - Started at {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
                _logFileWriter.WriteLine();
                
                // Zaktualizuj status
                if (LogFileStatusTextBlock != null)
                {
                    LogFileStatusTextBlock.Text = $"Zapis do: {Path.GetFileName(_logFilePath)}";
                    LogFileStatusTextBlock.Visibility = Visibility.Visible;
                }
                
                // Zapisz istniejące logi do pliku (jeśli są)
                if (LogTextBox != null && !string.IsNullOrEmpty(LogTextBox.Text))
                {
                    _logFileWriter.WriteLine("--- Existing logs ---");
                    _logFileWriter.Write(LogTextBox.Text);
                    _logFileWriter.WriteLine("--- End of existing logs ---");
                    _logFileWriter.WriteLine();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd podczas otwierania pliku logów: {ex.Message}");
                if (LogToFileCheckBox != null)
                {
                    LogToFileCheckBox.IsChecked = false;
                }
            }
        }
        
        public void StopLoggingToFile()
        {
            try
            {
                if (_logFileWriter != null)
                {
                    _logFileWriter.WriteLine($"=== Log ended at {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
                    _logFileWriter.Flush();
                    _logFileWriter.Close();
                    _logFileWriter.Dispose();
                    _logFileWriter = null;
                }
                
                _logFilePath = null;
                
                // Ukryj status
                if (LogFileStatusTextBlock != null)
                {
                    LogFileStatusTextBlock.Visibility = Visibility.Collapsed;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd podczas zamykania pliku logów: {ex.Message}");
            }
        }
        
        private void WriteToLogFile(string message)
        {
            if (_logFileWriter != null && _logFileWriter.BaseStream != null)
            {
                try
                {
                    _logFileWriter.Write(message);
                    _logFileWriter.Flush();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Błąd podczas zapisu do pliku logów: {ex.Message}");
                    // Wyłącz logowanie do pliku jeśli wystąpi błąd
                    if (LogToFileCheckBox != null)
                    {
                        LogToFileCheckBox.IsChecked = false;
                    }
                }
            }
        }
    }
}

