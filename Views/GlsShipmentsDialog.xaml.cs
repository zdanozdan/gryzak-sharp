using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Gryzak.Models;
using Gryzak.Services;
using static Gryzak.Services.Logger;

namespace Gryzak.Views
{
    public partial class GlsShipmentsDialog : Window
    {
        private const int PageSize = 250;

        private readonly GlsShipmentStore _store;
        private readonly ConfigService _configService;
        private readonly ObservableCollection<GlsShipmentRecord> _items = new();
        private int _totalCount;
        private bool _isLoading;
        private bool _hasMore = true;

        public GlsShipmentsDialog(GlsShipmentStore? store = null, ConfigService? configService = null)
        {
            InitializeComponent();
            _store = store ?? new GlsShipmentStore();
            _configService = configService ?? new ConfigService();
            DbPathText.Text = _store.DatabasePath;
            ShipmentsGrid.ItemsSource = _items;
            ShipmentsGrid.AddHandler(
                ScrollViewer.ScrollChangedEvent,
                new ScrollChangedEventHandler(ShipmentsGrid_ScrollChanged),
                handledEventsToo: true);
            Loaded += async (_, _) => await ReloadAsync();
        }

        private async void RefreshButton_Click(object sender, RoutedEventArgs e) =>
            await ReloadAsync();

        private void SearchGlsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new GlsParcelSearchDialog(_configService)
                {
                    Owner = this,
                    Title = "Szukaj w GLS"
                };
                dialog.ShowDialog();
            }
            catch (Exception ex)
            {
                Error(ex, "GlsShipmentsDialog", "Błąd otwierania wyszukiwania GLS");
                MessageBox.Show(
                    $"Nie udało się otworzyć wyszukiwania GLS.\n\n{ex.Message}",
                    "Szukaj w GLS",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async void ShipmentsGrid_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (!_hasMore || _isLoading || e.VerticalChange <= 0)
            {
                return;
            }

            var viewer = e.OriginalSource as ScrollViewer
                ?? FindVisualChild<ScrollViewer>(ShipmentsGrid);
            if (viewer == null || viewer.ScrollableHeight <= 0)
            {
                return;
            }

            if (viewer.VerticalOffset < viewer.ScrollableHeight - 80)
            {
                return;
            }

            await LoadMoreAsync();
        }

        private async System.Threading.Tasks.Task ReloadAsync()
        {
            if (_isLoading)
            {
                return;
            }

            _isLoading = true;
            RefreshButton.IsEnabled = false;
            StatusText.Text = "Ładowanie…";
            try
            {
                _items.Clear();
                _totalCount = await _store.CountAsync().ConfigureAwait(true);
                _hasMore = _totalCount > 0;
                if (_totalCount == 0)
                {
                    StatusText.Text = "Brak wpisów w cache gls_shipment.";
                    return;
                }

                await AppendPageAsync().ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                Error(ex, "GlsShipmentsDialog", "Błąd odczytu cache przesyłek");
                StatusText.Text = "Błąd odczytu cache.";
                MessageBox.Show(
                    $"Nie udało się wczytać tabeli gls_shipment.\n\n{ex.Message}",
                    "Przesyłki",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                _isLoading = false;
                RefreshButton.IsEnabled = true;
            }
        }

        private async System.Threading.Tasks.Task LoadMoreAsync()
        {
            if (_isLoading || !_hasMore)
            {
                return;
            }

            _isLoading = true;
            try
            {
                StatusText.Text = $"Ładowanie… ({_items.Count}/{_totalCount})";
                await AppendPageAsync().ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                Error(ex, "GlsShipmentsDialog", "Błąd doładowania cache przesyłek");
                StatusText.Text = $"Wczytano {_items.Count} z {_totalCount} (błąd doładowania).";
            }
            finally
            {
                _isLoading = false;
            }
        }

        private async System.Threading.Tasks.Task AppendPageAsync()
        {
            var page = await _store.GetPageAsync(_items.Count, PageSize).ConfigureAwait(true);
            foreach (var item in page)
            {
                _items.Add(item);
            }

            _hasMore = _items.Count < _totalCount && page.Count > 0;
            StatusText.Text = _hasMore
                ? $"Wczytano {_items.Count} z {_totalCount}"
                : $"Wpisów: {_totalCount}";
        }

        private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null)
            {
                return null;
            }

            for (var i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typed)
                {
                    return typed;
                }

                var nested = FindVisualChild<T>(child);
                if (nested != null)
                {
                    return nested;
                }
            }

            return null;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
    }
}
