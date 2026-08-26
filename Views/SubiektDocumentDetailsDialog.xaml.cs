using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using Gryzak.Models;
using Gryzak.Services;
using static Gryzak.Services.Logger;

namespace Gryzak.Views
{
    public partial class SubiektDocumentDetailsDialog : Window, INotifyPropertyChanged
    {
        private const string CreateWaybillButtonText = "Przygotowalnia GLS";
        private const string EditWaybillButtonText = "Edytuj przygotowalnię";
        private const string ViewWaybillButtonText = "Podgląd przesyłki";

        private SubiektDocument _document;
        private readonly ConfigService _configService;
        private readonly SubiektApiService _subiektApiService;
        private string _documentNumber = "";
        private string _originalNumber = "";
        private List<SubiektRelatedDocument> _relatedDocuments = new();
        private bool _isNavigatingRelated;
        private DateTime? _documentDate;
        private string _statusName = "";
        private string _customerName = "";
        private string _customerNip = "";
        private string _customerAddress = "";
        private string _payerName = "";
        private string _payerNip = "";
        private string _payerAddress = "";
        private bool _hasShippingAddress;
        private string _shippingAddressName = "";
        private string _shippingAddress = "";
        private decimal _totalNet;
        private decimal _totalGross;
        private string _payment = "";
        private List<SubiektDocumentLine> _lines = new();
        private bool _isSavingWaybill;
        private int? _fsDokId;
        private SubiektPrzesylka? _przesylka;
        private bool _glsWaybillExists;
        private string _waybillStatusText = "";
        private bool _isGlsLoading;
        private bool _glsListsLoaded;
        private List<GlsPreparingBoxItem> _preparingBoxItems = new();
        private List<GlsPickupItem> _pickupItems = new();

        public SubiektDocumentDetailsDialog(
            SubiektDocument document,
            ConfigService? configService = null,
            SubiektApiService? subiektApiService = null,
            bool glsPanelEnabled = false)
        {
            InitializeComponent();
            _document = document;
            _configService = configService ?? new ConfigService();
            _subiektApiService = subiektApiService ?? new SubiektApiService(_configService);
            ApplyDocument(document);
            DataContext = this;
            IsGlsPanelEnabled = glsPanelEnabled;
            if (!glsPanelEnabled)
            {
                GlsSection.Visibility = System.Windows.Visibility.Collapsed;
            }

            Loaded += OnLoaded;
            ContentHost.SizeChanged += (_, _) => QueueRelayoutContentCards();
            SizeChanged += (_, _) => QueueRelayoutContentCards();
        }

        public SubiektDocument Document => _document;

        public string DocumentNumber
        {
            get => _documentNumber;
            set { _documentNumber = value; OnPropertyChanged(); }
        }

        public string OriginalNumber
        {
            get => _originalNumber;
            set { _originalNumber = value; OnPropertyChanged(); }
        }

        public List<SubiektRelatedDocument> RelatedDocuments
        {
            get => _relatedDocuments;
            set
            {
                _relatedDocuments = value ?? new List<SubiektRelatedDocument>();
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasRelatedDocuments));
            }
        }

        public bool HasRelatedDocuments => RelatedDocuments.Count > 0;

        public DateTime? DocumentDate
        {
            get => _documentDate;
            set { _documentDate = value; OnPropertyChanged(); }
        }

        public string StatusName
        {
            get => _statusName;
            set { _statusName = value; OnPropertyChanged(); }
        }

        public string CustomerName
        {
            get => _customerName;
            set { _customerName = value; OnPropertyChanged(); }
        }

        public string CustomerNip
        {
            get => _customerNip;
            set { _customerNip = value; OnPropertyChanged(); }
        }

        public string CustomerAddress
        {
            get => _customerAddress;
            set { _customerAddress = value; OnPropertyChanged(); }
        }

        public string PayerName
        {
            get => _payerName;
            set { _payerName = value; OnPropertyChanged(); }
        }

        public string PayerNip
        {
            get => _payerNip;
            set { _payerNip = value; OnPropertyChanged(); }
        }

        public string PayerAddress
        {
            get => _payerAddress;
            set { _payerAddress = value; OnPropertyChanged(); }
        }

        public bool HasShippingAddress
        {
            get => _hasShippingAddress;
            set { _hasShippingAddress = value; OnPropertyChanged(); }
        }

        public string ShippingAddressName
        {
            get => _shippingAddressName;
            set { _shippingAddressName = value; OnPropertyChanged(); }
        }

        public string ShippingAddress
        {
            get => _shippingAddress;
            set { _shippingAddress = value; OnPropertyChanged(); }
        }

        public decimal TotalNet
        {
            get => _totalNet;
            set { _totalNet = value; OnPropertyChanged(); }
        }

        public decimal TotalGross
        {
            get => _totalGross;
            set { _totalGross = value; OnPropertyChanged(); }
        }

        public string Payment
        {
            get => _payment;
            set { _payment = value; OnPropertyChanged(); }
        }

        public bool IsGlsPanelEnabled { get; }

        private bool _relayoutQueued;

        public List<SubiektDocumentLine> Lines
        {
            get => _lines;
            set
            {
                _lines = value;
                OnPropertyChanged();
                QueueRelayoutContentCards();
            }
        }

        public string WaybillStatusText
        {
            get => _waybillStatusText;
            set { _waybillStatusText = value; OnPropertyChanged(); }
        }

        public bool IsGlsLoading
        {
            get => _isGlsLoading;
            set
            {
                if (_isGlsLoading == value) return;
                _isGlsLoading = value;
                OnPropertyChanged();
                NotifyGlsListVisibility();
            }
        }

        public bool GlsListsLoaded
        {
            get => _glsListsLoaded;
            set
            {
                if (_glsListsLoaded == value) return;
                _glsListsLoaded = value;
                OnPropertyChanged();
                NotifyGlsListVisibility();
            }
        }

        public List<GlsPreparingBoxItem> PreparingBoxItems
        {
            get => _preparingBoxItems;
            set
            {
                _preparingBoxItems = value ?? new List<GlsPreparingBoxItem>();
                OnPropertyChanged();
                NotifyGlsListVisibility();
            }
        }

        public List<GlsPickupItem> PickupItems
        {
            get => _pickupItems;
            set
            {
                _pickupItems = value ?? new List<GlsPickupItem>();
                OnPropertyChanged();
                NotifyGlsListVisibility();
            }
        }

        public bool ShowPreparingBoxList => !IsGlsLoading && PreparingBoxItems.Count > 0;
        public bool ShowPickupList => !IsGlsLoading && PickupItems.Count > 0;
        public bool ShowPreparingBoxEmpty => !IsGlsLoading && GlsListsLoaded && PreparingBoxItems.Count == 0;
        public bool ShowPickupEmpty => !IsGlsLoading && GlsListsLoaded && PickupItems.Count == 0;

        private void NotifyGlsListVisibility()
        {
            OnPropertyChanged(nameof(ShowPreparingBoxList));
            OnPropertyChanged(nameof(ShowPickupList));
            OnPropertyChanged(nameof(ShowPreparingBoxEmpty));
            OnPropertyChanged(nameof(ShowPickupEmpty));
            QueueRelayoutContentCards();
        }

        private void QueueRelayoutContentCards()
        {
            if (_relayoutQueued)
            {
                return;
            }

            _relayoutQueued = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                _relayoutQueued = false;
                RelayoutContentCards();
            }), DispatcherPriority.Loaded);
        }

        private void RelayoutContentCards()
        {
            if (ContentHost == null || ContentHost.ActualHeight <= 1)
            {
                return;
            }

            var available = ContentHost.ActualHeight;
            var width = ContentHost.ActualWidth;
            if (width <= 1)
            {
                return;
            }

            PreparingBoxScroll.MaxHeight = double.PositiveInfinity;
            PickupScroll.MaxHeight = double.PositiveInfinity;
            LinesScroll.MaxHeight = double.PositiveInfinity;

            var unconstrained = new Size(width, double.PositiveInfinity);
            GlsSection.Measure(unconstrained);
            ProductsSection.Measure(unconstrained);

            var glsDesired = GlsSection.DesiredSize.Height;
            var productsDesired = ProductsSection.DesiredSize.Height;
            if (glsDesired + productsDesired <= available)
            {
                return;
            }

            double glsAlloc;
            double productsAlloc;
            if (glsDesired <= available / 2)
            {
                glsAlloc = glsDesired;
                productsAlloc = available - glsDesired;
            }
            else if (productsDesired <= available / 2)
            {
                productsAlloc = productsDesired;
                glsAlloc = available - productsDesired;
            }
            else
            {
                glsAlloc = available / 2;
                productsAlloc = available - glsAlloc;
            }

            var columnWidth = Math.Max(0, (width - 48) / 2);
            var glsListNatural = Math.Max(
                MeasureScrollContent(PreparingBoxScroll, columnWidth),
                MeasureScrollContent(PickupScroll, columnWidth));
            var productsListNatural = MeasureScrollContent(LinesScroll, width);
            var glsChrome = Math.Max(0, glsDesired - glsListNatural);
            var productsChrome = Math.Max(0, productsDesired - productsListNatural);

            PreparingBoxScroll.MaxHeight = Math.Max(48, glsAlloc - glsChrome);
            PickupScroll.MaxHeight = PreparingBoxScroll.MaxHeight;
            LinesScroll.MaxHeight = Math.Max(48, productsAlloc - productsChrome);
        }

        private static double MeasureScrollContent(ScrollViewer viewer, double width)
        {
            if (viewer.Content is not UIElement content)
            {
                return 0;
            }

            content.Measure(new Size(Math.Max(width, 0), double.PositiveInfinity));
            return content.DesiredSize.Height;
        }

        private void ApplyDocument(SubiektDocument document)
        {
            Title = string.IsNullOrWhiteSpace(document.NrPelny)
                ? "Szczegóły dokumentu"
                : $"Szczegóły: {document.NrPelny}";

            DocumentNumber = document.NrPelny;
            OriginalNumber = document.NrPelnyOryg;
            RelatedDocuments = BuildRelatedFromDoDok(document);
            DocumentDate = document.DataWyst;
            StatusName = document.StatusNazwa;
            CustomerName = document.KontrahentNazwa;
            CustomerNip = document.KontrahentNip;
            CustomerAddress = document.KontrahentAdres;
            PayerName = document.PlatnikNazwa;
            PayerNip = document.PlatnikNip;
            PayerAddress = document.PlatnikAdres;
            HasShippingAddress = document.HasAdresDostawy;
            ShippingAddressName = document.AdresDostawyNazwa;
            ShippingAddress = document.AdresDostawyAdres;
            TotalNet = document.WartNetto;
            TotalGross = document.WartBrutto;
            Payment = document.PlatnoscDisplay;
            Lines = document.Pozycje ?? new List<SubiektDocumentLine>();
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Loaded -= OnLoaded;
            _przesylka = _document.Przesylka;
            _fsDokId = null;
            _glsWaybillExists = _przesylka is { HasId: true, IsDeleted: false };
            if (_przesylka != null && (!string.IsNullOrWhiteSpace(_przesylka.Status) || _przesylka.Data != null))
            {
                WaybillStatusText = _przesylka.SummaryText;
            }

            ShowSyncedStateWithoutGls();
            _ = LoadRelatedDocumentsAsync();
            _ = EnsurePrzesylkaForCardsAsync();
        }

        private async System.Threading.Tasks.Task EnsurePrzesylkaForCardsAsync()
        {
            var hadUsable = _przesylka is { IsDeleted: false } p
                && (p.HasId || !string.IsNullOrWhiteSpace(p.NrListuPrzygotowalnia) || !string.IsNullOrWhiteSpace(p.NrListuNadane));
            await EnsureFsDokIdAsync().ConfigureAwait(true);
            var hasUsable = _przesylka is { IsDeleted: false } q
                && (q.HasId || !string.IsNullOrWhiteSpace(q.NrListuPrzygotowalnia) || !string.IsNullOrWhiteSpace(q.NrListuNadane));
            if (!hadUsable && hasUsable)
            {
                ShowSyncedStateWithoutGls();
                QueueRelayoutContentCards();
            }
        }

        private async System.Threading.Tasks.Task LoadRelatedDocumentsAsync()
        {
            try
            {
                var items = await _subiektApiService.GetRelatedDocumentsExpandedAsync(
                    _document.DokId,
                    _document.TypKod).ConfigureAwait(true);

                RelatedDocuments = items.Count > 0
                    ? items
                    : BuildRelatedFromDoDok(_document);
            }
            catch (Exception ex)
            {
                Warning($"Nie udało się pobrać dokumentów powiązanych: {ex.Message}", "SubiektDocumentDetails");
            }
        }

        private static List<SubiektRelatedDocument> BuildRelatedFromDoDok(SubiektDocument document)
        {
            if (document.DoDokId is not int doDokId || doDokId <= 0)
            {
                return new List<SubiektRelatedDocument>();
            }

            var nr = SubiektDocumentNumber.Sanitize(document.DoDokNrPelny);
            if (string.IsNullOrWhiteSpace(nr))
            {
                return new List<SubiektRelatedDocument>();
            }

            return new List<SubiektRelatedDocument>
            {
                new()
                {
                    DokId = doDokId,
                    DokTyp = SubiektApiDocumentTypes.FromNrPelny(nr),
                    NrPelny = nr
                }
            };
        }

        private async void RelatedDocument_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is not FrameworkElement element
                || element.DataContext is not SubiektRelatedDocument related
                || related.DokId <= 0)
            {
                return;
            }

            e.Handled = true;
            await NavigateToRelatedDocumentAsync(related);
        }

        private async System.Threading.Tasks.Task NavigateToRelatedDocumentAsync(SubiektRelatedDocument related)
        {
            if (_isNavigatingRelated || related.DokId <= 0 || related.DokId == _document.DokId)
            {
                return;
            }

            _isNavigatingRelated = true;
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                var details = await _subiektApiService.GetDocumentByIdAsync(related.TypKod, related.DokId)
                    .ConfigureAwait(true);
                if (details == null)
                {
                    MessageBox.Show(
                        $"Nie udało się wczytać dokumentu {related.NrPelny}.",
                        "Dokument powiązany",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                _document = details;
                _przesylka = details.Przesylka;
                _fsDokId = null;
                _document.GlsPreparingBoxId = null;
                _document.GlsPreparingBoxParcelNumber = "";
                _document.GlsPickupConsignmentId = null;
                _document.GlsPickupParcelNumber = "";
                _document.GlsStatusChecked = false;
                GlsListsLoaded = false;
                ApplyDocument(details);

                // WZ/ZK nie mają pw_Przesylka — dociągnij z powiązanego FS (jak przy pierwszym otwarciu z listy).
                await EnsureFsDokIdAsync().ConfigureAwait(true);

                _glsWaybillExists = _przesylka is { HasId: true, IsDeleted: false };
                WaybillStatusText = _przesylka != null
                    && (!string.IsNullOrWhiteSpace(_przesylka.Status) || _przesylka.Data != null)
                        ? _przesylka.SummaryText
                        : "";
                ShowSyncedStateWithoutGls();
                await LoadRelatedDocumentsAsync().ConfigureAwait(true);
                QueueRelayoutContentCards();
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektDocumentDetails", "Błąd otwierania dokumentu powiązanego");
                MessageBox.Show(
                    $"Błąd otwierania dokumentu powiązanego: {ex.Message}",
                    "Błąd",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
                _isNavigatingRelated = false;
            }
        }

        private async System.Threading.Tasks.Task EnsureFsDokIdAsync()
        {
            var needsPrzesylka = _przesylka == null
                || _przesylka.IsDeleted
                || (!_przesylka.HasId && !_przesylka.HasAnyNrListu);

            if (_fsDokId is > 0 && !needsPrzesylka)
            {
                return;
            }

            try
            {
                var (fsId, przesylka) = await _subiektApiService.ResolveInvoiceShipmentAsync(_document);
                if (fsId is > 0)
                {
                    _fsDokId = fsId;
                }

                if (przesylka != null && needsPrzesylka)
                {
                    _przesylka = przesylka;
                    _document.Przesylka = przesylka;
                }
            }
            catch (Exception ex)
            {
                Warning($"Nie udało się odczytać pola Przesylka z FS: {ex.Message}", "SubiektDocumentDetails");
            }
        }

        private void ShowSyncedStateWithoutGls()
        {
            var przesylka = _przesylka ?? _document.Przesylka;
            var box = new List<GlsPreparingBoxItem>();
            var pickups = new List<GlsPickupItem>();

            if (_document.GlsPreparingBoxId is int boxId && boxId > 0)
            {
                box.Add(new GlsPreparingBoxItem
                {
                    Id = boxId,
                    References = _document.NrPelny,
                    ParcelNumber = _document.GlsPreparingBoxParcelNumber
                });
            }

            if (!string.IsNullOrWhiteSpace(_document.GlsPickupParcelNumber)
                || _document.GlsPickupConsignmentId is > 0)
            {
                pickups.Add(new GlsPickupItem
                {
                    Id = _document.GlsPickupConsignmentId ?? 0,
                    References = _document.NrPelny,
                    ParcelNumber = _document.GlsPickupParcelNumber
                });
            }

            // Jak kolumna Nadania na liście: gdy brak dopasowania GLS po referencji WZ/ZK,
            // numery często są na powiązanym FS (Przesylka) — pokaż je w kartach.
            if (przesylka is { IsDeleted: false })
            {
                if (box.Count == 0 && przesylka.HasId)
                {
                    box.Add(new GlsPreparingBoxItem
                    {
                        Id = przesylka.Id,
                        References = _document.NrPelny,
                        ParcelNumber = przesylka.NrListuPrzygotowalnia
                    });
                    if (_document.GlsPreparingBoxId is not > 0)
                    {
                        _document.GlsPreparingBoxId = przesylka.Id;
                        _document.GlsPreparingBoxParcelNumber =
                            SubiektPrzesylka.NormalizeNrListu(przesylka.NrListuPrzygotowalnia);
                    }
                }

                if (pickups.Count == 0 && !string.IsNullOrWhiteSpace(przesylka.NrListuNadane))
                {
                    pickups.Add(new GlsPickupItem
                    {
                        References = _document.NrPelny,
                        ParcelNumber = przesylka.NrListuNadane
                    });
                    if (string.IsNullOrWhiteSpace(_document.GlsPickupParcelNumber))
                    {
                        _document.GlsPickupParcelNumber =
                            SubiektPrzesylka.NormalizeNrListu(przesylka.NrListuNadane);
                    }
                }
            }

            PreparingBoxItems = box;
            PickupItems = pickups;
            _glsWaybillExists = box.Count > 0;
            GlsListsLoaded = true;
            WaybillStatusText = przesylka?.SummaryText ?? "";
            UpdateWaybillButton();
        }

        private async System.Threading.Tasks.Task<bool> RefreshGlsAndSyncAsync()
        {
            var config = _configService.LoadGlsConfig();
            if (string.IsNullOrWhiteSpace(config.UserName) || string.IsNullOrWhiteSpace(config.Password))
            {
                GlsListsLoaded = true;
                PreparingBoxItems = new List<GlsPreparingBoxItem>();
                PickupItems = new List<GlsPickupItem>();
                _glsWaybillExists = _przesylka is { HasId: true, IsDeleted: false };
                WaybillStatusText = _przesylka is { HasId: true }
                    ? $"Subiekt: id {_przesylka.Id} (nie sprawdzono w GLS — brak loginu)."
                    : "Nie sprawdzono GLS — brak loginu.";
                UpdateWaybillButton();
                return false;
            }

            IsGlsLoading = true;
            WaybillButton.IsEnabled = false;
            try
            {
                using var gls = new GlsService(config);
                var result = await gls.GetPreparingBoxListAsync();
                if (!result.Success)
                {
                    GlsListsLoaded = true;
                    PreparingBoxItems = new List<GlsPreparingBoxItem>();
                    PickupItems = new List<GlsPickupItem>();
                    WaybillStatusText = string.IsNullOrWhiteSpace(result.ErrorMessage)
                        ? "Nie udało się pobrać przygotowalni GLS."
                        : result.ErrorMessage;
                    Warning(WaybillStatusText, "SubiektDocumentDetails");
                    UpdateWaybillButton();
                    return false;
                }

                var box = GlsSubiektSync.MatchPreparingBox(_document, result.Items);
                var pickups = new List<GlsPickupItem>();

                var przesylka = _przesylka ?? _document.Przesylka;
                if (box.Count == 0
                    && przesylka is { IsDeleted: false, HasId: true } liveBox
                    && result.Items.Any(item => item.Id == liveBox.Id))
                {
                    box = result.Items.Where(item => item.Id == liveBox.Id).ToList();
                }

                if (przesylka is { IsDeleted: false }
                    && !string.IsNullOrWhiteSpace(przesylka.NrListuNadane))
                {
                    pickups.Add(new GlsPickupItem
                    {
                        References = _document.NrPelny,
                        ParcelNumber = przesylka.NrListuNadane
                    });
                }

                PreparingBoxItems = box;
                PickupItems = pickups;

                _document.GlsPreparingBoxId = box.FirstOrDefault()?.Id;
                _document.GlsPreparingBoxParcelNumber = GlsSubiektSync.CombinedParcelNumbers(box);
                _document.GlsPickupParcelNumber = GlsSubiektSync.CombinedParcelNumbers(pickups);
                _document.GlsPickupConsignmentId = pickups.FirstOrDefault()?.Id;
                _document.GlsStatusChecked = true;
                _glsWaybillExists = box.Count > 0;

                var liveIds = result.Items
                    .Select(item => item.Id)
                    .Where(id => id > 0)
                    .ToHashSet();
                try
                {
                    var sync = await GlsSubiektSync.ApplyToSubiektAsync(
                        _subiektApiService,
                        _document,
                        liveIds,
                        _document.GlsPreparingBoxId,
                        _document.GlsPreparingBoxParcelNumber,
                        matchedPickupNrListu: null);
                    if (sync.FsDokId is int fsDokId)
                    {
                        _fsDokId = fsDokId;
                    }

                    if (sync.Przesylka != null)
                    {
                        _przesylka = sync.Przesylka;
                    }
                }
                catch (Exception ex)
                {
                    Warning($"Nie udało się zsynchronizować Przesylka z GLS: {ex.Message}", "SubiektDocumentDetails");
                }

                WaybillStatusText = BuildGlsStatusText(box, pickups, pickupErrorMessage: null);
                GlsListsLoaded = true;
                UpdateWaybillButton();
                return true;
            }
            catch (Exception ex)
            {
                GlsListsLoaded = true;
                WaybillStatusText = $"Nie udało się sprawdzić GLS: {ex.Message}";
                Warning(WaybillStatusText, "SubiektDocumentDetails");
                UpdateWaybillButton();
                return false;
            }
            finally
            {
                IsGlsLoading = false;
                WaybillButton.IsEnabled = !_isSavingWaybill;
            }
        }

        private static string BuildGlsStatusText(
            IReadOnlyCollection<GlsPreparingBoxItem> box,
            IReadOnlyCollection<GlsPickupItem> pickups,
            string? pickupErrorMessage)
        {
            var parts = new List<string>();
            if (box.Count == 0 && pickups.Count == 0)
            {
                parts.Add("Brak przesyłek w przygotowalni GLS.");
            }
            else
            {
                if (box.Count > 0)
                {
                    parts.Add($"Przygotowalnia: {box.Count}");
                }

                if (pickups.Count > 0)
                {
                    parts.Add($"Nr listu (Subiekt): {pickups.Count}");
                }
            }

            if (!string.IsNullOrWhiteSpace(pickupErrorMessage))
            {
                parts.Add(pickupErrorMessage);
            }

            return string.Join("  •  ", parts);
        }

        private void UpdateWaybillButton()
        {
            WaybillButton.Content = HasIssuedLabel()
                ? ViewWaybillButtonText
                : (HasExistingWaybill ? EditWaybillButtonText : CreateWaybillButtonText);
        }

        private bool HasExistingWaybill => _glsWaybillExists;

        private async void CreateWaybillButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isSavingWaybill)
            {
                return;
            }

            var config = _configService.LoadGlsConfig();
            if (string.IsNullOrWhiteSpace(config.UserName) || string.IsNullOrWhiteSpace(config.Password))
            {
                MessageBox.Show(
                    "Brak loginu lub hasła GLS.\n\nUzupełnij je w menu Ustawienia → Ustawienia GLS.",
                    "Brak konfiguracji GLS",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            await EnsureFsDokIdAsync();

            var consignment = GlsConsignment.FromDocument(_document);
            var editId = GetPreparingBoxEditId();
            if (HasExistingWaybill && editId is > 0)
            {
                consignment.ExistingId = editId;
                Mouse.OverrideCursor = Cursors.Wait;
                WaybillButton.IsEnabled = false;
                try
                {
                    using var glsLookup = new GlsService(config);
                    var existing = await glsLookup.GetShipmentAsync(editId.Value);
                    if (existing.Success && existing.Consignment != null)
                    {
                        var codAmountFallback = consignment.CodAmount;
                        var forceCod = consignment.CashOnDelivery;
                        consignment = existing.Consignment;
                        consignment.ExistingId = editId;
                        if (consignment.CodAmount <= 0 && codAmountFallback > 0)
                        {
                            consignment.CodAmount = codAmountFallback;
                        }

                        if (forceCod)
                        {
                            consignment.CashOnDelivery = true;
                        }
                    }
                    else if (existing.NotFound)
                    {
                        _glsWaybillExists = false;
                        consignment.ExistingId = null;
                        WaybillStatusText =
                            $"W Subiekcie jest id {editId}, ale w GLS tej przesyłki już nie ma.";
                        UpdateWaybillButton();
                        Info(WaybillStatusText, "SubiektDocumentDetails");
                    }
                    else
                    {
                        Warning(
                        existing.ErrorMessage ?? "Nie udało się pobrać przesyłki z przygotowalni GLS — pokazano dane z dokumentu.",
                        "SubiektDocumentDetails");
                    }
                }
                catch (Exception ex)
                {
                    Warning($"Nie udało się pobrać przesyłki z przygotowalni GLS: {ex.Message}", "SubiektDocumentDetails");
                }
                finally
                {
                    Mouse.OverrideCursor = null;
                    WaybillButton.IsEnabled = true;
                }
            }

            var form = new GlsWaybillDialog(
                consignment,
                config,
                existingNrListu: GetExistingNrListu(),
                configService: _configService)
            {
                Owner = this
            };

            var dialogResult = form.ShowDialog();

            // Po etykiecie numery listu są już w GLS — zapisz i pokaż na karcie bez ręcznego Odśwież.
            if (form.LabelIssued && !form.DeleteRequested)
            {
                await ApplyIssuedLabelAsync(
                    form.Consignment.ExistingId ?? GetPreparingBoxEditId(),
                    form.IssuedParcelNumbers);
            }

            if (dialogResult != true)
            {
                return;
            }

            if (form.DeleteRequested)
            {
                await DeleteWaybillAsync(config, form.Consignment.ExistingId ?? GetPreparingBoxEditId());
                return;
            }

            consignment = form.Consignment;
            var isEdit = HasExistingWaybill && consignment.ExistingId is > 0;
            _isSavingWaybill = true;
            WaybillButton.IsEnabled = false;
            WaybillButton.Content = isEdit ? "Zapisywanie..." : "Dodawanie...";
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                using var glsService = new GlsService(config);
                var result = isEdit
                    ? await glsService.UpdateShipmentAsync(consignment.ExistingId!.Value, consignment)
                    : await glsService.CreateShipmentAsync(consignment);

                if (!result.Success)
                {
                    var error = result.ErrorMessage ?? "Nieznany błąd GLS.";
                    Warning(error, "SubiektDocumentDetails");
                    MessageBox.Show(
                        $"Nie udało się {(isEdit ? "zapisać" : "dodać")} przesyłki w przygotowalni GLS ({result.EnvironmentName}).\n\n{error}",
                        "Błąd GLS",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                var savedPrzesylka = await SavePrzesylkaToSubiektAsync(result.ConsignmentId!.Value, isEdit);
                await RefreshGlsAndSyncAsync();
                var message = isEdit
                    ? $"Przesyłka GLS ({result.EnvironmentName}) została zapisana w przygotowalni.\n\nIdentyfikator: {result.ConsignmentId}"
                    : $"Przesyłka GLS ({result.EnvironmentName}) została dodana do przygotowalni.\n\nIdentyfikator: {result.ConsignmentId}";

                if (!string.IsNullOrWhiteSpace(result.WarningMessage))
                {
                    message += "\n\n" + result.WarningMessage;
                }

                if (!string.IsNullOrWhiteSpace(savedPrzesylka.Warning))
                {
                    message += "\n\n" + savedPrzesylka.Warning;
                }

                Info(message, "SubiektDocumentDetails");
                MessageBox.Show(
                    message,
                    isEdit ? "Przygotowalnia GLS — zapisano" : "Przygotowalnia GLS — dodano",
                    MessageBoxButton.OK,
                    string.IsNullOrWhiteSpace(result.WarningMessage) && string.IsNullOrWhiteSpace(savedPrzesylka.Warning)
                        ? MessageBoxImage.Information
                        : MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektDocumentDetails", "Błąd przygotowalni GLS");
                MessageBox.Show($"Błąd przygotowalni GLS:\n\n{ex.Message}", "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _isSavingWaybill = false;
                Mouse.OverrideCursor = null;
                UpdateWaybillButton();
                WaybillButton.IsEnabled = true;
            }
        }

        private int? GetPreparingBoxEditId()
        {
            if (_przesylka is { HasId: true })
            {
                var matching = PreparingBoxItems.FirstOrDefault(item => item.Id == _przesylka.Id);
                if (matching != null)
                {
                    return matching.Id;
                }
            }

            return PreparingBoxItems.FirstOrDefault()?.Id
                ?? _document.GlsPreparingBoxId
                ?? (_przesylka is { HasId: true } ? _przesylka.Id : null);
        }

        private string GetExistingNrListu()
        {
            return SubiektPrzesylka.NormalizeNrListu(
                GlsSubiektSync.CombineNrListu(
                    _document.GlsPreparingBoxParcelNumber,
                    _document.Przesylka?.NrListuPrzygotowalnia,
                    _przesylka?.NrListuPrzygotowalnia));
        }

        private bool HasIssuedLabel() => !string.IsNullOrWhiteSpace(GetExistingNrListu());

        /// <summary>
        /// Po adePreparingBox_GetConsignLabels GLS nadaje numery paczek — zapisujemy je
        /// w Przesylka i od razu na karcie przygotowalni (bez czekania na pełne Odśwież).
        /// </summary>
        private async System.Threading.Tasks.Task ApplyIssuedLabelAsync(
            int? consignmentId,
            string? issuedParcelNumbers)
        {
            if (consignmentId is not > 0)
            {
                return;
            }

            var nr = SubiektPrzesylka.NormalizeNrListu(issuedParcelNumbers);
            if (string.IsNullOrWhiteSpace(nr))
            {
                // Numery jeszcze nie wróciły z GetConsign — pełna synchronizacja listy GLS.
                await RefreshGlsAndSyncAsync();
                return;
            }

            _document.GlsPreparingBoxId = consignmentId;
            _document.GlsPreparingBoxParcelNumber = nr;
            _document.GlsStatusChecked = true;

            var existingNadane = SubiektPrzesylka.NormalizeNrListu(
                _przesylka?.NrListuNadane
                ?? _document.Przesylka?.NrListuNadane
                ?? _document.GlsPickupParcelNumber);

            var next = new SubiektPrzesylka
            {
                Typ = "GLS",
                Id = consignmentId.Value,
                NrListuPrzygotowalnia = nr,
                NrListuNadane = existingNadane,
                Status = SubiektPrzesylka.StatusEdytowano,
                Data = DateTime.Now
            };

            await EnsureFsDokIdAsync();
            if (_fsDokId is > 0)
            {
                try
                {
                    var saved = await _subiektApiService.PutPrzesylkaAsync(_fsDokId.Value, next);
                    _przesylka = saved ?? next;
                    _document.Przesylka = _przesylka;
                }
                catch (Exception ex)
                {
                    Warning(
                        $"Etykieta GLS ma numery {nr}, ale nie zapisano Przesylka na FS: {ex.Message}",
                        "SubiektDocumentDetails");
                    _przesylka = next;
                    _document.Przesylka = next;
                }
            }
            else
            {
                _przesylka = next;
                _document.Przesylka = next;
            }

            var box = new List<GlsPreparingBoxItem>
            {
                new()
                {
                    Id = consignmentId.Value,
                    References = _document.NrPelny,
                    ParcelNumber = nr
                }
            };
            PreparingBoxItems = box;
            _glsWaybillExists = true;

            if (PickupItems.Count == 0 && !string.IsNullOrWhiteSpace(existingNadane))
            {
                PickupItems = new List<GlsPickupItem>
                {
                    new()
                    {
                        References = _document.NrPelny,
                        ParcelNumber = existingNadane
                    }
                };
                _document.GlsPickupParcelNumber = existingNadane;
            }

            WaybillStatusText = BuildGlsStatusText(PreparingBoxItems, PickupItems, pickupErrorMessage: null);
            GlsListsLoaded = true;
            UpdateWaybillButton();
            Info(
                $"Po etykiecie GLS zapisano nr listu przygotowalni: {nr} (id={consignmentId}).",
                "SubiektDocumentDetails");
        }

        private async System.Threading.Tasks.Task<(SubiektPrzesylka? Przesylka, string? Warning)> SavePrzesylkaToSubiektAsync(
            int consignmentId,
            bool isEdit)
        {
                var existingNadane = SubiektPrzesylka.NormalizeNrListu(
                    _przesylka?.NrListuNadane
                    ?? _document.Przesylka?.NrListuNadane);
                var next = new SubiektPrzesylka
                {
                    Typ = "GLS",
                    Id = consignmentId,
                    NrListuPrzygotowalnia = isEdit
                        ? (!string.IsNullOrWhiteSpace(_przesylka?.NrListuPrzygotowalnia) ? _przesylka!.NrListuPrzygotowalnia : "")
                        : "",
                    NrListuNadane = existingNadane,
                    Status = isEdit ? SubiektPrzesylka.StatusEdytowano : SubiektPrzesylka.StatusUtworzono,
                    Data = DateTime.Now
                };

            if (_fsDokId is not > 0)
            {
                var warning = "Przesyłka GLS została dodana do przygotowalni, ale nie zapisano jej na fakturze — brak powiązanego FS z polem Przesylka.";
                Warning(warning, "SubiektDocumentDetails");
                _przesylka = next;
                _document.Przesylka = next;
                return (next, warning);
            }

            try
            {
                var saved = await _subiektApiService.PutPrzesylkaAsync(_fsDokId.Value, next);
                _przesylka = saved ?? next;
                _document.Przesylka = _przesylka;
                return (_przesylka, null);
            }
            catch (Exception ex)
            {
                var warning = $"Przesyłka GLS w przygotowalni ma id={consignmentId}, ale nie udało się zapisać pola Przesylka na FS: {ex.Message}";
                Warning(warning, "SubiektDocumentDetails");
                _przesylka = next;
                _document.Przesylka = next;
                return (next, warning);
            }
        }

        private async System.Threading.Tasks.Task DeleteWaybillAsync(GlsConfig config, int? consignmentId)
        {
            if (consignmentId is not > 0)
            {
                MessageBox.Show("Brak identyfikatora przesyłki do usunięcia.", "Usuń list", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _isSavingWaybill = true;
            WaybillButton.IsEnabled = false;
            WaybillButton.Content = "Usuwanie...";
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                using var glsService = new GlsService(config);
                var deleted = await glsService.DeleteShipmentAsync(consignmentId.Value);
                if (!deleted.Success && !deleted.NotFound)
                {
                    var error = deleted.ErrorMessage ?? "Nieznany błąd usuwania przesyłki GLS.";
                    Warning(error, "SubiektDocumentDetails");
                    MessageBox.Show(
                        $"Nie udało się usunąć przesyłki z przygotowalni GLS ({deleted.EnvironmentName}).\n\nWpis w Subiekcie nie został zmieniony.\n\n{error}",
                        "Błąd GLS",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                string? subiektWarning = null;
                var refreshed = await RefreshGlsAndSyncAsync();
                if (!refreshed && _fsDokId is > 0)
                {
                    try
                    {
                        var cleared = await _subiektApiService.ClearPrzesylkaAsync(_fsDokId.Value);
                        _przesylka = cleared ?? new SubiektPrzesylka
                        {
                            Typ = "GLS",
                            Status = SubiektPrzesylka.StatusUsuniete,
                            Data = DateTime.Now
                        };
                        _document.Przesylka = _przesylka;
                        _glsWaybillExists = false;
                    }
                    catch (Exception ex)
                    {
                        subiektWarning = $"List usunięto z GLS, ale nie udało się zaktualizować pola Przesylka na FS: {ex.Message}";
                        Warning(subiektWarning, "SubiektDocumentDetails");
                    }
                }

                string message;
                if (_fsDokId is not > 0)
                {
                    message = $"Przesyłka GLS ({deleted.EnvironmentName}) została usunięta z przygotowalni.\n\nBrak powiązanego FS — pole Przesylka nie było aktualizowane.";
                }
                else if (!string.IsNullOrWhiteSpace(subiektWarning))
                {
                    message = $"Przesyłkę usunięto z przygotowalni GLS ({deleted.EnvironmentName}).\n\n{subiektWarning}";
                }
                else if (deleted.NotFound)
                {
                    message = "Przesyłki nie było już w przygotowalni GLS. Pole Przesylka na fakturze zsynchronizowano z GLS.";
                }
                else
                {
                    message = $"Przesyłka GLS ({deleted.EnvironmentName}) została usunięta z przygotowalni. Pole Przesylka na fakturze zsynchronizowano z GLS.";
                }

                Info(message, "SubiektDocumentDetails");
                MessageBox.Show(
                    message,
                    "Usuń z przygotowalni",
                    MessageBoxButton.OK,
                    string.IsNullOrWhiteSpace(subiektWarning) ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektDocumentDetails", "Błąd usuwania z przygotowalni GLS");
                MessageBox.Show($"Błąd usuwania z przygotowalni GLS:\n\n{ex.Message}", "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _isSavingWaybill = false;
                Mouse.OverrideCursor = null;
                UpdateWaybillButton();
                WaybillButton.IsEnabled = true;
            }
        }

        private async void PayerPanel_Click(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            var khId = _document.PlatnikId;
            if (khId <= 0)
            {
                MessageBox.Show(
                    "Brak identyfikatora płatnika na dokumencie.",
                    "Płatnik",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                var details = await _subiektApiService.GetKontrahentDetailsAsync(khId);
                if (details == null)
                {
                    MessageBox.Show(
                        $"Nie znaleziono kontrahenta (kh_Id={khId}).",
                        "Płatnik",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                var dialog = new KontrahentDetailsDialog(details, "Płatnik")
                {
                    Owner = this
                };
                dialog.ShowDialog();
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektDocumentDetails", "Błąd pobierania danych płatnika");
                MessageBox.Show(
                    $"Nie udało się pobrać danych płatnika:\n\n{ex.Message}",
                    "Błąd",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
