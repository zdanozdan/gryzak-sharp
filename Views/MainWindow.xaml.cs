using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Gryzak.Models;
using Gryzak.ViewModels;

namespace Gryzak.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            SetupKeyBindings();
            SetupStatusFilter();
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

        private void OpenOrderDetailsDialog(Order order, MainViewModel vm)
        {
            var detailsDialog = new OrderDetailsDialog(order, vm)
            {
                Owner = this
            };
            detailsDialog.ShowDialog();
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
                        Console.WriteLine($"[Gryzak] Scroll blisko końca, ładuję kolejną stronę");
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
                Application.Current.Shutdown();
            }
        }

        private void AboutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var aboutDialog = new AboutDialog
            {
                Owner = this
            };
            aboutDialog.ShowDialog();
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

