using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Gryzak.Models;
using Gryzak.ViewModels;
using static Gryzak.Services.Logger;

namespace Gryzak.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;
            this.Closing += MainWindow_Closing;
        }
        
        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            // Zwolnij licencję Subiekta GT przed zamknięciem aplikacji
            // Musimy to zrobić synchronicznie, aby upewnić się że licencja jest zwolniona przed zamknięciem
            try
            {
                var subiektService = new Gryzak.Services.SubiektService();
                subiektService.ZwolnijLicencje();
                Info("Licencja Subiekta GT zwolniona przed zamknięciem okna.", "MainWindow");
            }
            catch (Exception ex)
            {
                Error(ex, "MainWindow", "Błąd podczas zwalniania licencji przy zamykaniu");
            }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            SetupKeyBindings();
            SetupStatusFilter();
            SetupActivityTracking();
        }
        
        private void SetupActivityTracking()
        {
            // Nasłuchuj aktywności na całym oknie
            this.MouseMove += (s, e) => ResetActivityTimer();
            this.KeyDown += (s, e) => ResetActivityTimer();
            this.MouseDown += (s, e) => ResetActivityTimer();
            
            // Nasłuchuj aktywności w kontrolkach
            if (SearchTextBox != null)
            {
                SearchTextBox.TextChanged += (s, e) => ResetActivityTimer();
                SearchTextBox.KeyDown += (s, e) => ResetActivityTimer();
                SearchTextBox.MouseMove += (s, e) => ResetActivityTimer();
            }
            
            if (StatusFilterComboBox != null)
            {
                StatusFilterComboBox.SelectionChanged += (s, e) => ResetActivityTimer();
                StatusFilterComboBox.MouseMove += (s, e) => ResetActivityTimer();
            }
        }
        
        private void ResetActivityTimer()
        {
            if (DataContext is MainViewModel vm)
            {
                vm.ResetActivityTimer();
            }
        }
        
        private void Window_MouseMove(object sender, MouseEventArgs e)
        {
            ResetActivityTimer();
        }
        
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            ResetActivityTimer();
        }

        private void SetupStatusFilter()
        {
            if (DataContext is MainViewModel vm)
            {
                // Ustaw domyślną wartość w ComboBox
                if (StatusFilterComboBox.Items.Count > 0)
                {
                    StatusFilterComboBox.SelectedIndex = 0;
                    vm.StatusFilter = "Wszystkie statusy";
                }
            }
        }

        private void SetupKeyBindings()
        {
            // Ctrl+K - Konfiguracja API
            if (DataContext is MainViewModel vm)
            {
                var configBinding = new KeyBinding(
                    vm.ConfigureApiCommand,
                    Key.K,
                    ModifierKeys.Control);
                this.InputBindings.Add(configBinding);
            }

            // Ctrl+Q - Zamknij
            var closeBinding = new KeyBinding(
                new RelayCommand(() => CloseMenuItem_Click(null!, null!)),
                Key.Q,
                ModifierKeys.Control);
            this.InputBindings.Add(closeBinding);
        }

        private void StatusFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DataContext is MainViewModel vm && sender is ComboBox comboBox)
            {
                if (comboBox.SelectedItem is ComboBoxItem item)
                {
                    vm.StatusFilter = item.Content?.ToString() ?? "Wszystkie statusy";
                }
            }
        }

        private void OrderItem_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is Order order)
            {
                if (DataContext is MainViewModel vm)
                {
                    // Sprawdź czy to podwójne kliknięcie
                    if (e.ClickCount == 2)
                    {
                        // Podwójne kliknięcie - zaznacz zamówienie i otwórz okno szczegółów
                        vm.OrderSelectedCommand.Execute(order);
                        OpenOrderDetailsDialog(order, vm);
                    }
                    else if (e.ClickCount == 1)
                    {
                        // Pojedyncze kliknięcie - tylko zaznacz zamówienie
                        vm.OrderSelectedCommand.Execute(order);
                    }
                }
            }
        }

        public void OpenOrderDetailsDialog(Order order, MainViewModel vm)
        {
            if (order == null)
            {
                System.Diagnostics.Debug.WriteLine("OpenOrderDetailsDialog: order jest null, przerywam");
                return;
            }
            
            System.Diagnostics.Debug.WriteLine($"OpenOrderDetailsDialog wywołana dla zamówienia: {order.Id}");
            var detailsDialog = new OrderDetailsDialog(order, vm)
            {
                Owner = this
            };
            System.Diagnostics.Debug.WriteLine("Tworzenie OrderDetailsDialog zakończone, wywołuję ShowDialog");
            detailsDialog.ShowDialog();
            System.Diagnostics.Debug.WriteLine("ShowDialog zakończone");
        }

        private async void OrdersScrollViewer_ScrollChanged(object sender, System.Windows.Controls.ScrollChangedEventArgs e)
        {
            if (sender is ScrollViewer scrollViewer && DataContext is MainViewModel vm)
            {
                // Jeśli trwa wyszukiwanie (niepusty SearchText), nie uruchamiaj infinite scroll
                if (!string.IsNullOrWhiteSpace(vm.SearchText))
                {
                    return;
                }

                // Sprawdź czy użytkownik jest blisko końca (500px od końca)
                var scrollOffset = scrollViewer.VerticalOffset;
                var scrollHeight = scrollViewer.ScrollableHeight;
                var viewportHeight = scrollViewer.ViewportHeight;

                // Jeśli użytkownik jest blisko końca (500px lub mniej)
                if (scrollHeight - scrollOffset - viewportHeight <= 500)
                {
                    if (vm.HasMorePages && !vm.IsLoadingMore && !vm.IsLoading)
                    {
                        Debug("Scroll blisko końca, ładuję kolejną stronę", "MainWindow");
                        await vm.LoadNextPageAsync();
                    }
                }
            }
        }

        private void CloseMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "Czy na pewno chcesz zamknąć aplikację?",
                "Potwierdzenie",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                // Licencja zostanie zwolniona w MainWindow_Closing
                Application.Current.Shutdown();
            }
        }

        private void DebugMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var debugWindow = DebugWindow.GetOrCreateInstance();
            
            // Okno jest tworzone w osobnym wątku, więc jest całkowicie niezależne
            // od modalnych okien (ShowDialog nie zablokuje tego okna)
            debugWindow.ShowWindow();
        }

        private void AboutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var aboutDialog = new AboutDialog
            {
                Owner = this
            };
            aboutDialog.ShowDialog();
        }

        private void ReadmeMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var readmeDialog = new ReadmeDialog
            {
                Owner = this
            };
            readmeDialog.ShowDialog();
        }

        private void HistoryOrderItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem && menuItem.DataContext is string orderId)
            {
                System.Diagnostics.Debug.WriteLine($"HistoryOrderItem_Click: Kliknięto zamówienie {orderId}");
                if (DataContext is MainViewModel vm)
                {
                    System.Diagnostics.Debug.WriteLine($"HistoryOrderItem_Click: Wywołuję OpenOrderFromHistory dla {orderId}");
                    vm.OpenOrderFromHistoryCommand.Execute(orderId);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("HistoryOrderItem_Click: DataContext nie jest MainViewModel");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"HistoryOrderItem_Click: sender nie jest MenuItem lub DataContext nie jest string. sender={sender?.GetType().Name}, DataContext={((sender as MenuItem)?.DataContext)?.GetType().Name}");
            }
        }

        private void LogoImage_MouseDown(object sender, MouseButtonEventArgs e)
        {
            // Otwórz okno z dużym logo
            var logoWindow = new Window
            {
                Title = "Gryzak Logo",
                Width = 400,
                Height = 400,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                Background = System.Windows.Media.Brushes.White,
                ResizeMode = ResizeMode.NoResize
            };

            var image = new System.Windows.Controls.Image
            {
                Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("pack://application:,,,/gryzak.png")),
                Stretch = System.Windows.Media.Stretch.Uniform,
                Margin = new Thickness(20)
            };

            logoWindow.Content = image;
            logoWindow.ShowDialog();
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action _execute;

        public RelayCommand(Action execute)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        }

        public event EventHandler? CanExecuteChanged
        {
            add { }
            remove { }
        }

        public bool CanExecute(object? parameter) => true;

        public void Execute(object? parameter) => _execute();
    }
}

