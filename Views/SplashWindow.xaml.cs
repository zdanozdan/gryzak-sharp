using System;
using System.Windows;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace Gryzak.Views
{
    public partial class SplashWindow : Window
    {
        private readonly DispatcherTimer _progressTimer;
        private readonly DispatcherTimer _closeTimer;
        private double _progress = 0;

        public SplashWindow()
        {
            InitializeComponent();
            
            // Timer do animacji postępu
            _progressTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(50)
            };
            _progressTimer.Tick += ProgressTimer_Tick;

            // Timer do zamykania okna
            _closeTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(4)
            };
            _closeTimer.Tick += CloseTimer_Tick;

            Loaded += SplashWindow_Loaded;
        }

        private void SplashWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Wyłącz automatyczne timery - będziemy kontrolować postęp ręcznie
            // _progressTimer.Start();
            // StartLoadingSequence();
            // _closeTimer.Start();
            
            // Ustaw początkowy stan
            ProgressBar.Value = 0;
            LoadingText.Text = "Inicjalizacja...";
        }

        private void ProgressTimer_Tick(object? sender, EventArgs e)
        {
            _progress += 1.67; // 100% w 3 sekundy (3000ms / 50ms * 0.017 = 100%)
            
            if (_progress >= 100)
            {
                _progress = 100;
                _progressTimer.Stop();
            }

            ProgressBar.Value = _progress;
        }

        private void StartLoadingSequence()
        {
            var sequence = new DispatcherTimer[]
            {
                new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500), Tag = "Inicjalizacja..." },
                new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(1200), Tag = "Ładowanie danych..." },
                new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(2200), Tag = "Przygotowywanie interfejsu..." },
                new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(2800), Tag = "Finalizacja..." },
                new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(3500), Tag = "Gotowe!" }
            };

            foreach (var timer in sequence)
            {
                timer.Tick += (s, e) =>
                {
                    LoadingText.Text = timer.Tag?.ToString() ?? "Ładowanie...";
                    ((DispatcherTimer)s!).Stop();
                };
                timer.Start();
            }
        }

        private void CloseTimer_Tick(object? sender, EventArgs e)
        {
            _closeTimer.Stop();
            Close();
        }

        public void CloseSplash()
        {
            _progressTimer?.Stop();
            _closeTimer?.Stop();
            Close();
        }
        
        // Metody do aktualizacji postępu z zewnątrz
        public void UpdateProgress(double value, string text)
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBar.Value = Math.Min(100, Math.Max(0, value));
                LoadingText.Text = text;
            });
        }
        
        public void UpdateText(string text)
        {
            Dispatcher.Invoke(() =>
            {
                LoadingText.Text = text;
            });
        }
    }
}

