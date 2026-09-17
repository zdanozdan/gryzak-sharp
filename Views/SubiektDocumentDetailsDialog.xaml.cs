using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using Gryzak.Models;
using Gryzak.Services;
using Gryzak.Services.Llm;
using static Gryzak.Services.Logger;

namespace Gryzak.Views
{
    public partial class SubiektDocumentDetailsDialog : Window, INotifyPropertyChanged
    {
        private const string CreateWaybillButtonText = "Przygotowalnia GLS";
        private const string EditWaybillButtonText = "Edytuj przygotowalnię";
        private const string ViewWaybillButtonText = "Podgląd przesyłki";
        private const string NonShopShippingNote =
            "Brak adresu ze sklepu. Możesz wyodrębnić adres dostawy z uwag przyciskiem AI.";

        private SubiektDocument _document;
        private readonly ConfigService _configService;
        private readonly SubiektApiService _subiektApiService;
        private readonly ApiService _apiService;
        private readonly GlsShipmentStore _glsShipmentStore = new();
        private readonly DocumentAiAddressStore _aiAddressStore = new();
        private CancellationTokenSource? _shippingLoadCts;
        private CancellationTokenSource? _aiExtractCts;
        private Task _shippingLoadTask = Task.CompletedTask;
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
        private string _payerNameShort = "";
        private string _payerNip = "";
        private string _payerAddress = "";
        private string _payerEmail = "";
        private string _payerPhone = "";
        private bool _hasSubiektShippingAddress;
        private string _subiektShippingName = "";
        private string _subiektShippingAddress = "";
        private string _subiektShippingPhone = "";
        private string _documentUwagi = "";
        private bool _hasShippingAddress;
        private string _shippingAddressName = "";
        private string _shippingAddress = "";
        private string _shippingPersonName = "";
        private string _shippingEmail = "";
        private string _shippingPhone = "";
        private string _shippingSourceSubtitle = "";
        private string _shippingNote = "";
        private bool _isAiExtracting;
        private string _aiExtractStatusText = "";
        private bool _hasAiShippingUi;
        private string _aiShippingSourceSubtitle = "";
        private string _aiShippingPersonName = "";
        private string _aiShippingAddress = "";
        private string _aiShippingPhone = "";
        private string _aiShippingEmail = "";
        private string _aiShippingNote = "";
        private decimal _totalNet;
        private decimal _totalGross;
        private string _payment = "";
        private List<SubiektDocumentLine> _lines = new();
        private bool _isSavingWaybill;
        private int? _przesylkaDokId;
        private string _przesylkaTypKod = "";
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
            _apiService = new ApiService(_configService);
            ApplyDocument(document);
            DataContext = this;
            IsGlsPanelEnabled = glsPanelEnabled;
            if (!glsPanelEnabled)
            {
                GlsSection.Visibility = System.Windows.Visibility.Collapsed;
            }

            Loaded += OnLoaded;
            Closed += (_, _) =>
            {
                _shippingLoadCts?.Cancel();
                _aiExtractCts?.Cancel();
            };
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

        public string PayerNameShort
        {
            get => _payerNameShort;
            set { _payerNameShort = value; OnPropertyChanged(); }
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

        public string PayerEmail
        {
            get => _payerEmail;
            set { _payerEmail = value; OnPropertyChanged(); }
        }

        public string PayerPhone
        {
            get => _payerPhone;
            set { _payerPhone = value; OnPropertyChanged(); }
        }

        public bool HasSubiektShippingAddress
        {
            get => _hasSubiektShippingAddress;
            set { _hasSubiektShippingAddress = value; OnPropertyChanged(); }
        }

        public string SubiektShippingName
        {
            get => _subiektShippingName;
            set { _subiektShippingName = value; OnPropertyChanged(); }
        }

        public string SubiektShippingAddress
        {
            get => _subiektShippingAddress;
            set { _subiektShippingAddress = value; OnPropertyChanged(); }
        }

        public string SubiektShippingPhone
        {
            get => _subiektShippingPhone;
            set { _subiektShippingPhone = value; OnPropertyChanged(); }
        }

        public string DocumentUwagi
        {
            get => _documentUwagi;
            set
            {
                _documentUwagi = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanExtractAiAddress));
            }
        }

        public bool CanExtractAiAddress =>
            !_isAiExtracting && !string.IsNullOrWhiteSpace(DocumentUwagi);

        public bool CanNormalizeShopShippingAi =>
            !_isAiExtracting
            && (_document.HasShopShippingAddress
                || !string.IsNullOrWhiteSpace(ShippingPersonName)
                || !string.IsNullOrWhiteSpace(ShippingAddressName)
                || !string.IsNullOrWhiteSpace(ShippingAddress));

        public bool IsAiExtracting
        {
            get => _isAiExtracting;
            set
            {
                if (_isAiExtracting == value) return;
                _isAiExtracting = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanExtractAiAddress));
                OnPropertyChanged(nameof(CanNormalizeShopShippingAi));
            }
        }

        public string AiExtractStatusText
        {
            get => _aiExtractStatusText;
            set { _aiExtractStatusText = value; OnPropertyChanged(); }
        }

        public bool HasAiShippingUi
        {
            get => _hasAiShippingUi;
            set { _hasAiShippingUi = value; OnPropertyChanged(); }
        }

        public string AiShippingSourceSubtitle
        {
            get => _aiShippingSourceSubtitle;
            set { _aiShippingSourceSubtitle = value; OnPropertyChanged(); }
        }

        public string AiShippingPersonName
        {
            get => _aiShippingPersonName;
            set { _aiShippingPersonName = value; OnPropertyChanged(); }
        }

        public string AiShippingAddress
        {
            get => _aiShippingAddress;
            set { _aiShippingAddress = value; OnPropertyChanged(); }
        }

        public string AiShippingPhone
        {
            get => _aiShippingPhone;
            set { _aiShippingPhone = value; OnPropertyChanged(); }
        }

        public string AiShippingEmail
        {
            get => _aiShippingEmail;
            set { _aiShippingEmail = value; OnPropertyChanged(); }
        }

        public string AiShippingNote
        {
            get => _aiShippingNote;
            set { _aiShippingNote = value; OnPropertyChanged(); }
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

        public string ShippingPersonName
        {
            get => _shippingPersonName;
            set { _shippingPersonName = value; OnPropertyChanged(); }
        }

        public string ShippingEmail
        {
            get => _shippingEmail;
            set { _shippingEmail = value; OnPropertyChanged(); }
        }

        public string ShippingPhone
        {
            get => _shippingPhone;
            set { _shippingPhone = value; OnPropertyChanged(); }
        }

        public string ShippingSourceSubtitle
        {
            get => _shippingSourceSubtitle;
            set { _shippingSourceSubtitle = value; OnPropertyChanged(); }
        }

        public string ShippingNote
        {
            get => _shippingNote;
            set { _shippingNote = value; OnPropertyChanged(); }
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
            PayerNameShort = document.PlatnikNazwaKrotka;
            PayerNip = document.PlatnikNip;
            PayerAddress = document.PlatnikAdres;
            PayerEmail = document.PlatnikEmail;
            PayerPhone = document.PlatnikTelefon;
            HasSubiektShippingAddress = document.HasAdresDostawy;
            SubiektShippingName = document.HasAdresDostawy ? document.AdresDostawyNazwa : "";
            SubiektShippingAddress = document.HasAdresDostawy ? document.AdresDostawyAdres : "";
            SubiektShippingPhone = document.HasAdresDostawy ? document.AdresDostawyTelefon : "";
            DocumentUwagi = document.Uwagi ?? "";
            TotalNet = document.WartNetto;
            TotalGross = document.WartBrutto;
            Payment = document.PlatnoscDisplay;
            Lines = document.Pozycje ?? new List<SubiektDocumentLine>();
            _shippingLoadTask = LoadShippingAddressAsync();
        }

        private async Task LoadShippingAddressAsync()
        {
            _shippingLoadCts?.Cancel();
            _shippingLoadCts?.Dispose();
            _shippingLoadCts = new CancellationTokenSource();
            var ct = _shippingLoadCts.Token;
            var dokId = _document.DokId;

            ClearShippingUi();
            ClearAiShippingUi();
            _document.ClearShopShippingAddress();
            _document.ClearAiShippingAddress();
            _document.ShopOrderId = "";
            AiExtractStatusText = "";
            HasShippingAddress = true;

            if (!SubiektDocument.TryResolveShopOrderId(_document, out var orderId))
            {
                // Na WZ/FS uwagi i numer oryginalny często nie przechodzą z ZK — szukaj w powiązanych.
                var fromRelated = await TryResolveShopOrderIdFromRelatedAsync(ct).ConfigureAwait(true);
                if (ct.IsCancellationRequested || _document.DokId != dokId)
                    return;
                orderId = fromRelated.OrderId ?? "";
                if (string.IsNullOrWhiteSpace(DocumentUwagi) && !string.IsNullOrWhiteSpace(fromRelated.Uwagi))
                {
                    DocumentUwagi = fromRelated.Uwagi;
                    _document.Uwagi = fromRelated.Uwagi;
                }
                if (string.IsNullOrWhiteSpace(orderId))
                {
                    ShippingNote = NonShopShippingNote;
                    await TryApplyCachedAiAddressAsync(ct).ConfigureAwait(true);
                    return;
                }
            }

            _document.ShopOrderId = orderId;
            ShippingSourceSubtitle = $"Sklep #{orderId}";
            ShippingAddress = "Ładowanie adresu…";

            try
            {
                var json = await _apiService.GetOrderDetailsAsync(orderId, ct).ConfigureAwait(true);
                if (ct.IsCancellationRequested || _document.DokId != dokId)
                    return;

                using var parsed = JsonDocument.Parse(json);
                var root = parsed.RootElement;
                var order = new Order();
                order.ApplyAddressesFromApi(root);

                var email = ReadOrderString(root, "email");
                var phone = ReadOrderString(root, "telephone");
                if (string.IsNullOrWhiteSpace(phone))
                    phone = ReadOrderString(root, "shipping_telephone");
                var country = ReadOrderString(root, "shipping_iso_code_2");
                if (string.IsNullOrWhiteSpace(country))
                    country = ReadOrderString(root, "payment_iso_code_2");

                if (!string.IsNullOrWhiteSpace(email) && email != "Brak email")
                    order.Email = email;
                if (!string.IsNullOrWhiteSpace(phone) && phone != "Brak telefonu")
                    order.Phone = phone;

                if (!order.HasShippingAddress)
                {
                    ClearShippingUi();
                    HasShippingAddress = true;
                    ShippingSourceSubtitle = $"Sklep #{orderId}";
                    ShippingNote = "Brak adresu wysyłki w zamówieniu sklepu.";
                    await TryApplyCachedAiAddressAsync(ct).ConfigureAwait(true);
                    return;
                }

                ApplyShopShippingToDocument(order, country);
                ApplyShopShippingToUi(order, orderId);
                await TryApplyCachedAiAddressAsync(ct).ConfigureAwait(true);
            }
            catch (OperationCanceledException)
            {
                // nawigacja do innego dokumentu / zamknięcie okna
            }
            catch (Exception ex)
            {
                if (ct.IsCancellationRequested || _document.DokId != dokId)
                    return;

                Warning($"Nie udało się pobrać adresu wysyłki dla zamówienia {orderId}: {ex.Message}",
                    "SubiektDocumentDetails");
                ClearShippingUi();
                HasShippingAddress = true;
                ShippingSourceSubtitle = $"Sklep #{orderId}";
                _document.ClearShopShippingAddress();
                ShippingNote = $"Nie udało się pobrać adresu ze sklepu: {ex.Message}";
                await TryApplyCachedAiAddressAsync(ct).ConfigureAwait(true);
                // ShopOrderId zostaje — do pola uwagi GLS na WZ/FS.
            }
        }

        /// <summary>
        /// Szuka numeru zamówienia sklepu na dokumentach powiązanych (np. ZK przy WZ/FS).
        /// W GT uwagi i <c>dok_NrPelnyOryg</c> często nie przenoszą się z ZK na WZ.
        /// </summary>
        private async Task<(string? OrderId, string? Uwagi)> TryResolveShopOrderIdFromRelatedAsync(
            CancellationToken cancellationToken)
        {
            List<SubiektRelatedDocument> related;
            try
            {
                related = await _subiektApiService.GetRelatedDocumentsExpandedAsync(
                    _document.DokId,
                    _document.TypKod,
                    cancellationToken: cancellationToken).ConfigureAwait(true);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Warning(
                    $"Nie udało się pobrać powiązanych przy szukaniu nr zamówienia: {ex.Message}",
                    "SubiektDocumentDetails");
                related = new List<SubiektRelatedDocument>();
            }

            if (related.Count == 0)
                related = BuildRelatedFromDoDok(_document);

            // ZK ma numer zamówienia / uwagi; potem FS, WZ, reszta.
            var candidates = related
                .Where(r => r.DokId > 0 && r.DokId != _document.DokId)
                .OrderBy(r => r.IsZk ? 0 : r.IsFs ? 1 : r.IsWz ? 2 : 3)
                .ThenBy(r => r.DokId)
                .ToList();

            string? uwagiFromRelated = null;
            string? orderId = null;

            foreach (var relatedDoc in candidates)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var details = await _subiektApiService
                        .GetDocumentByIdAsync(relatedDoc.TypKod, relatedDoc.DokId, cancellationToken: cancellationToken)
                        .ConfigureAwait(true);
                    if (details == null)
                        continue;

                    if (string.IsNullOrWhiteSpace(uwagiFromRelated) && !string.IsNullOrWhiteSpace(details.Uwagi))
                        uwagiFromRelated = details.Uwagi;

                    if (orderId == null && SubiektDocument.TryResolveShopOrderId(details, out var foundId))
                    {
                        Debug(
                            $"Nr zamówienia sklepu {foundId} z powiązanego {relatedDoc.NrPelny} (bieżący {_document.NrPelny})",
                            "SubiektDocumentDetails");
                        orderId = foundId;
                        // ZK zwykle ma uwagi — nie trzeba skanować dalej, jeśli mamy i nr, i uwagi.
                        if (!string.IsNullOrWhiteSpace(uwagiFromRelated) || relatedDoc.IsZk)
                            return (orderId, uwagiFromRelated ?? details.Uwagi);
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    Warning(
                        $"Nie udało się odczytać {relatedDoc.NrPelny} przy szukaniu nr zamówienia: {ex.Message}",
                        "SubiektDocumentDetails");
                }
            }

            return (orderId, uwagiFromRelated);
        }

        private void ClearShippingUi()
        {
            ShippingAddressName = "";
            ShippingAddress = "";
            ShippingPersonName = "";
            ShippingEmail = "";
            ShippingPhone = "";
            ShippingSourceSubtitle = "";
            ShippingNote = "";
            OnPropertyChanged(nameof(CanNormalizeShopShippingAi));
        }

        private void ClearAiShippingUi()
        {
            HasAiShippingUi = false;
            AiShippingSourceSubtitle = "";
            AiShippingPersonName = "";
            AiShippingAddress = "";
            AiShippingPhone = "";
            AiShippingEmail = "";
            AiShippingNote = "";
        }

        private void ApplyShopShippingToDocument(Order order, string? countryCode)
        {
            _document.HasShopShippingAddress = true;
            var person = order.ShippingPersonName;
            if (string.IsNullOrWhiteSpace(person))
                person = order.Customer?.Trim() ?? "";
            _document.ShopShippingPersonName = person;
            _document.ShopShippingCompany = order.ShippingCompanyDisplay;
            _document.ShopShippingStreet = order.ShippingDisplayStreet;
            _document.ShopShippingKodPocztowy = order.ShippingPostcode?.Trim() ?? "";
            _document.ShopShippingMiejscowosc = order.ShippingCity?.Trim() ?? "";
            _document.ShopShippingEmail = NormalizeContact(order.Email, "Brak email");
            _document.ShopShippingTelefon = NormalizeContact(order.Phone, "Brak telefonu");
            _document.ShopShippingKrajKod = (countryCode ?? "").Trim();
        }

        private void ApplyShopShippingToUi(Order order, string orderId)
        {
            ShippingSourceSubtitle = $"Sklep #{orderId}";
            ShippingAddressName = order.ShippingCompanyDisplay;

            // Imię i nazwisko — z shipping_*, albo z danych płatności / klienta zamówienia.
            var person = order.ShippingPersonName;
            if (string.IsNullOrWhiteSpace(person))
                person = order.Customer?.Trim() ?? "";
            ShippingPersonName = person;

            var addressParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(order.ShippingDisplayStreet))
                addressParts.Add(order.ShippingDisplayStreet);
            if (!string.IsNullOrWhiteSpace(order.ShippingDisplayCityLine))
                addressParts.Add(order.ShippingDisplayCityLine);
            ShippingAddress = string.Join("\n", addressParts);

            ShippingEmail = NormalizeContact(order.Email, "Brak email");
            ShippingPhone = NormalizeContact(order.Phone, "Brak telefonu");
            ShippingNote = "";
            HasShippingAddress = true;
            OnPropertyChanged(nameof(CanNormalizeShopShippingAi));
        }

        private async Task<bool> TryApplyCachedAiAddressAsync(CancellationToken cancellationToken)
        {
            try
            {
                var related = RelatedDocuments?.Count > 0
                    ? RelatedDocuments
                    : BuildRelatedFromDoDok(_document);
                var (record, _, _) = await _aiAddressStore
                    .GetForDocumentAsync(_document, related, cancellationToken)
                    .ConfigureAwait(true);
                if (record == null || record.IsEmpty)
                {
                    return false;
                }

                ApplyAiAddressToDocumentAndUi(record, fromCache: true, sourceTag: "#ai");
                return true;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Warning($"Nie udało się odczytać adresu AI z SQLite: {ex.Message}", "SubiektDocumentDetails");
                return false;
            }
        }

        private void ApplyAiAddressToDocumentAndUi(
            DocumentAiAddress address,
            bool fromCache,
            string sourceTag = "#ai z uwag")
        {
            void Apply()
            {
                _document.ApplyAiShippingAddress(address);

                var confidence = string.IsNullOrWhiteSpace(address.Confidence)
                    ? ""
                    : $" · {address.Confidence}";
                AiShippingSourceSubtitle = fromCache
                    ? $"{sourceTag}{confidence}"
                    : $"{sourceTag} (nowy){confidence}";
                AiShippingPersonName = address.FormatDisplayName();
                AiShippingAddress = address.FormatDisplayAddress();
                AiShippingPhone = address.Phone ?? "";
                AiShippingEmail = address.Contact ?? "";
                AiShippingNote = string.IsNullOrWhiteSpace(address.Notes) ? "" : address.Notes;
                // Wymuś odświeżenie Visibility (false→true), gdy blok był już widoczny.
                if (_hasAiShippingUi)
                {
                    _hasAiShippingUi = false;
                    OnPropertyChanged(nameof(HasAiShippingUi));
                }

                HasAiShippingUi = true;

                AiExtractStatusText = fromCache
                    ? "Adres AI z lokalnej bazy."
                    : "Adres AI zapisany lokalnie (pierwszeństwo przy GLS).";
            }

            if (Dispatcher.CheckAccess())
            {
                Apply();
            }
            else
            {
                Dispatcher.Invoke(Apply);
            }
        }

        private async void DeleteAiAddressButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "Usunąć zapisany adres AI z lokalnej bazy dla tego dokumentu i powiązanych?",
                "Usuń adres AI",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
            {
                return;
            }

            try
            {
                var related = await ResolveRelatedForAiAsync(CancellationToken.None).ConfigureAwait(true);
                await _aiAddressStore
                    .DeleteForDocumentAndRelatedAsync(_document, related, CancellationToken.None)
                    .ConfigureAwait(true);

                _document.ClearAiShippingAddress();
                ClearAiShippingUi();
                AiExtractStatusText = "Usunięto adres AI z lokalnej bazy.";
                Info($"Usunięto adres AI dla {_document.NrPelny} i powiązanych.", "SubiektDocumentDetails");
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektDocumentDetails", "Błąd usuwania adresu AI");
                MessageBox.Show(
                    $"Nie udało się usunąć adresu AI:\n{ex.Message}",
                    "Usuń adres AI",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async void ExtractAiAddressButton_Click(object sender, RoutedEventArgs e)
        {
            var uwagi = DocumentUwagi?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(uwagi))
            {
                MessageBox.Show(
                    "Brak tekstu uwag do analizy.",
                    "AI — adres GLS",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            await RunAiAddressActionAsync(
                    statusText: "Analizuję uwagi przez AI…",
                    sourceTag: "#ai z uwag",
                    sourceHashText: uwagi,
                    invoke: (service, ct) => service.ExtractAsync(uwagi, ct))
                .ConfigureAwait(true);
        }

        private async void NormalizeShopAddressAiButton_Click(object sender, RoutedEventArgs e)
        {
            var shopText = BuildShopShippingTextForAi();
            if (string.IsNullOrWhiteSpace(shopText))
            {
                MessageBox.Show(
                    "Brak adresu wysyłki ze sklepu do normalizacji.",
                    "AI — adres GLS",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            await RunAiAddressActionAsync(
                    statusText: "Normalizuję adres sklepu przez AI…",
                    sourceTag: "#ai ze sklepu",
                    sourceHashText: shopText,
                    invoke: (service, ct) => service.NormalizeShopAddressAsync(shopText, ct))
                .ConfigureAwait(true);
        }

        private string BuildShopShippingTextForAi()
        {
            var lines = new List<string>();
            void Add(string label, string? value)
            {
                if (!string.IsNullOrWhiteSpace(value))
                    lines.Add($"{label}: {value.Trim()}");
            }

            if (_document.HasShopShippingAddress)
            {
                Add("firma", _document.ShopShippingCompany);
                Add("osoba", _document.ShopShippingPersonName);
                Add("ulica", _document.ShopShippingStreet);
                Add("kod", _document.ShopShippingKodPocztowy);
                Add("miasto", _document.ShopShippingMiejscowosc);
                Add("kraj", _document.ShopShippingKrajKod);
                Add("telefon", _document.ShopShippingTelefon);
                Add("email", _document.ShopShippingEmail);
                Add("nr_zamowienia", _document.ShopOrderId);
            }
            else
            {
                Add("osoba", ShippingPersonName);
                Add("firma", ShippingAddressName);
                Add("adres", ShippingAddress);
                Add("telefon", ShippingPhone);
                Add("email", ShippingEmail);
                if (!string.IsNullOrWhiteSpace(ShippingSourceSubtitle))
                    lines.Add($"zrodlo: {ShippingSourceSubtitle}");
            }

            return string.Join("\n", lines);
        }

        private async Task RunAiAddressActionAsync(
            string statusText,
            string sourceTag,
            string sourceHashText,
            Func<GlsAddressExtractionService, CancellationToken, Task<GlsAddressExtractionResult>> invoke)
        {
            var aiConfig = _configService.LoadAiConfig();
            if (string.IsNullOrWhiteSpace(aiConfig.ApiKey))
            {
                MessageBox.Show(
                    "Brak klucza API w Ustawienia → AI.",
                    "AI — adres GLS",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            _aiExtractCts?.Cancel();
            _aiExtractCts?.Dispose();
            _aiExtractCts = new CancellationTokenSource();
            var ct = _aiExtractCts.Token;
            var dokId = _document.DokId;

            IsAiExtracting = true;
            AiExtractStatusText = statusText;

            try
            {
                using var client = LlmClientFactory.Create(aiConfig);
                var service = new GlsAddressExtractionService(client);
                var result = await invoke(service, ct).ConfigureAwait(true);
                if (ct.IsCancellationRequested || _document.DokId != dokId)
                    return;

                if (!result.Success)
                {
                    AiExtractStatusText = result.ErrorMessage ?? "Błąd AI.";
                    MessageBox.Show(
                        result.ErrorMessage ?? "Nie udało się wywołać AI.",
                        "AI — adres GLS",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                if (!result.Found || result.Address == null)
                {
                    AiExtractStatusText = result.ErrorMessage ?? "AI nie zwróciło kompletnego adresu.";
                    MessageBox.Show(
                        result.ErrorMessage ?? "Model nie zwrócił kompletnego adresu dostawy.",
                        "AI — adres GLS",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                var address = result.Address;
                address.DokId = _document.DokId;
                address.DokTyp = _document.DokTyp;
                address.NrPelny = _document.NrPelny ?? "";
                address.UwagiHash = GlsAddressExtractionService.ComputeUwagiHash(sourceHashText);

                await _aiAddressStore.UpsertAsync(address, ct).ConfigureAwait(true);

                var related = await ResolveRelatedForAiAsync(ct).ConfigureAwait(true);
                await _aiAddressStore.PropagateToRelatedAsync(address, related, ct).ConfigureAwait(true);

                if (ct.IsCancellationRequested || _document.DokId != dokId)
                    return;

                ApplyAiAddressToDocumentAndUi(address, fromCache: false, sourceTag: sourceTag);
                Info(
                    $"AI adres GLS ({sourceTag}) dla {_document.NrPelny}: {address.FormatDisplayName()}, {address.Street}, {address.ZipCode} {address.City}",
                    "SubiektDocumentDetails");
            }
            catch (OperationCanceledException)
            {
                // zamknięcie / kolejny test
            }
            catch (Exception ex)
            {
                if (ct.IsCancellationRequested || _document.DokId != dokId)
                    return;
                AiExtractStatusText = ex.Message;
                Error(ex, "SubiektDocumentDetails", "Błąd AI adresu");
                MessageBox.Show(
                    $"Błąd AI:\n{ex.Message}",
                    "AI — adres GLS",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                IsAiExtracting = false;
            }
        }

        private async Task<List<SubiektRelatedDocument>> ResolveRelatedForAiAsync(
            CancellationToken cancellationToken)
        {
            var list = new List<SubiektRelatedDocument>();
            void Add(SubiektRelatedDocument? doc)
            {
                if (doc == null || doc.DokId <= 0)
                    return;
                if (list.Any(x => x.DokId == doc.DokId && x.DokTyp == doc.DokTyp))
                    return;
                list.Add(doc);
            }

            Add(new SubiektRelatedDocument
            {
                DokId = _document.DokId,
                DokTyp = _document.DokTyp,
                NrPelny = _document.NrPelny ?? ""
            });

            foreach (var item in RelatedDocuments ?? new List<SubiektRelatedDocument>())
                Add(item);

            foreach (var item in BuildRelatedFromDoDok(_document))
                Add(item);

            try
            {
                var expanded = await _subiektApiService.GetRelatedDocumentsExpandedAsync(
                    _document.DokId,
                    _document.TypKod,
                    cancellationToken: cancellationToken).ConfigureAwait(true);
                foreach (var item in expanded)
                    Add(item);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Warning(
                    $"Nie udało się pobrać powiązanych przy zapisie adresu AI: {ex.Message}",
                    "SubiektDocumentDetails");
            }

            return list;
        }

        private static string NormalizeContact(string? value, string placeholder)
        {
            var text = (value ?? "").Trim();
            if (string.IsNullOrEmpty(text) || text == placeholder)
                return "";
            return text;
        }

        private static string? ReadOrderString(JsonElement root, string name)
        {
            if (!root.TryGetProperty(name, out var prop))
                return null;
            if (prop.ValueKind == JsonValueKind.String)
                return prop.GetString()?.Trim();
            if (prop.ValueKind == JsonValueKind.Number)
                return prop.ToString();
            return null;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Loaded -= OnLoaded;
            _przesylka = _document.Przesylka;
            if ((_przesylka == null || _przesylka.IsEmpty)
                && _document.GlsShipment is { IsEmpty: false } cached)
            {
                ApplyShipmentToPrzesylka(cached);
            }

            _przesylkaDokId = _document.PrzesylkaOwnerDokId
                ?? _document.GlsShipmentDokId
                ?? _document.DokId;
            _przesylkaTypKod = !string.IsNullOrWhiteSpace(_document.PrzesylkaOwnerTypKod)
                ? _document.PrzesylkaOwnerTypKod
                : (_document.GlsShipmentDokTyp is int typ
                    ? SubiektApiDocumentTypes.FromDokTyp(typ)
                    : SubiektApiDocumentTypes.FromDokTyp(_document.DokTyp));
            _glsWaybillExists = _przesylka is { HasId: true };
            if (_przesylka != null && (_przesylka.HasId || _przesylka.HasAnyNrListu))
            {
                WaybillStatusText = _przesylka.SummaryText;
            }

            ShowSyncedStateWithoutGls();
            _ = LoadRelatedDocumentsAsync();
            _ = EnsurePrzesylkaForCardsAsync();
        }

        private async System.Threading.Tasks.Task EnsurePrzesylkaForCardsAsync()
        {
            var hadUsable = _przesylka is { IsEmpty: false };
            await EnsurePrzesylkaOwnerAsync().ConfigureAwait(true);
            var hasUsable = _przesylka is { IsEmpty: false };
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

                // Po pełnej liście powiązań — dociągnij zapisany adres AI, jeśli jeszcze go nie ma w UI.
                if (!HasAiShippingUi)
                {
                    await TryApplyCachedAiAddressAsync(CancellationToken.None).ConfigureAwait(true);
                }
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
                _przesylkaDokId = details.PrzesylkaOwnerDokId;
                _przesylkaTypKod = details.PrzesylkaOwnerTypKod ?? "";
                _document.GlsPreparingBoxId = null;
                _document.GlsPreparingBoxParcelNumber = "";
                _document.GlsPickupConsignmentId = null;
                _document.GlsPickupParcelNumber = "";
                _document.GlsStatusChecked = false;
                GlsListsLoaded = false;
                ApplyDocument(details);

                // Na FS/ZK bez własnego pola — dociągnij Przesylka z powiązanego WZ (lub legacy FS).
                await EnsurePrzesylkaOwnerAsync().ConfigureAwait(true);

                _glsWaybillExists = _przesylka is { HasId: true };
                WaybillStatusText = _przesylka != null
                    && (_przesylka.HasId || _przesylka.HasAnyNrListu)
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

        private async System.Threading.Tasks.Task EnsurePrzesylkaOwnerAsync(bool forWrite = false)
        {
            var needsPrzesylka = _przesylka == null || _przesylka.IsEmpty;

            if (_przesylkaDokId is > 0
                && !string.IsNullOrWhiteSpace(_przesylkaTypKod)
                && !needsPrzesylka
                && !forWrite)
            {
                return;
            }

            try
            {
                if (RelatedDocuments.Count == 0)
                {
                    await LoadRelatedDocumentsAsync().ConfigureAwait(true);
                }

                var (record, sourceDokId, sourceDokTyp) = await _glsShipmentStore
                    .GetForDocumentAsync(_document, RelatedDocuments)
                    .ConfigureAwait(true);

                if (record is { IsEmpty: false })
                {
                    _przesylkaDokId = sourceDokId;
                    _przesylkaTypKod = SubiektApiDocumentTypes.FromDokTyp(sourceDokTyp);
                    _document.PrzesylkaOwnerDokId = sourceDokId;
                    _document.PrzesylkaOwnerTypKod = _przesylkaTypKod;
                    GlsSubiektSync.ApplyCacheToDocument(_document, record);
                    if (needsPrzesylka || forWrite)
                    {
                        ApplyShipmentToPrzesylka(record);
                    }

                    return;
                }

                if (forWrite || _przesylkaDokId is not > 0)
                {
                    _przesylkaDokId = _document.DokId;
                    _przesylkaTypKod = SubiektApiDocumentTypes.FromDokTyp(_document.DokTyp);
                    if (string.IsNullOrWhiteSpace(_przesylkaTypKod))
                    {
                        _przesylkaTypKod = SubiektApiService.NormalizeDocumentType(_document.TypKod);
                    }

                    _document.PrzesylkaOwnerDokId = _przesylkaDokId;
                    _document.PrzesylkaOwnerTypKod = _przesylkaTypKod;
                }
            }
            catch (Exception ex)
            {
                Warning($"Nie udało się odczytać lokalnego cache GLS: {ex.Message}", "SubiektDocumentDetails");
            }
        }

        private void ApplyShipmentToPrzesylka(GlsShipmentRecord? shipment)
        {
            if (shipment == null || shipment.IsEmpty)
            {
                _przesylka = null;
                _document.Przesylka = null;
                return;
            }

            _przesylka = new SubiektPrzesylka
            {
                Typ = string.IsNullOrWhiteSpace(shipment.Carrier) ? "GLS" : shipment.Carrier,
                Id = shipment.BoxId,
                NrListuPrzygotowalnia = shipment.NrPrzyg,
                NrListuNadane = shipment.NrNad
            };
            _document.Przesylka = _przesylka;
        }

        private void ApplyLocalSyncResult(GlsLocalSyncResult sync)
        {
            if (sync.CacheDokId is int ownerDokId)
            {
                _przesylkaDokId = ownerDokId;
                _przesylkaTypKod = sync.CacheTypKod;
                _document.PrzesylkaOwnerDokId = ownerDokId;
                _document.PrzesylkaOwnerTypKod = sync.CacheTypKod;
                _document.GlsShipmentDokId = ownerDokId;
                _document.GlsShipmentDokTyp = sync.CacheDokTyp;
            }

            ApplyShipmentToPrzesylka(sync.Shipment);
        }

        private async System.Threading.Tasks.Task ApplyLocalSyncResultAndPropagateAsync(GlsLocalSyncResult sync)
        {
            ApplyLocalSyncResult(sync);

            try
            {
                if (RelatedDocuments.Count == 0)
                {
                    await LoadRelatedDocumentsAsync().ConfigureAwait(true);
                }

                var sourceDokId = sync.CacheDokId ?? _document.DokId;
                var sourceDokTyp = sync.CacheDokTyp > 0 ? sync.CacheDokTyp : _document.DokTyp;
                await _glsShipmentStore
                    .PropagateToRelatedAsync(
                        sync.Shipment,
                        sourceDokId,
                        sourceDokTyp,
                        RelatedDocuments)
                    .ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                Warning(
                    $"Nie udało się spropagować cache GLS na dokumenty powiązane: {ex.Message}",
                    "SubiektDocumentDetails");
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

            // Gdy brak dopasowania GLS po referencji — numery z pola Przesylka (WZ/FS, ewentualnie related).
            if (przesylka is { IsEmpty: false })
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
                _glsWaybillExists = _przesylka is { HasId: true };
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

                var cacheIndex = GlsShipmentCacheIndex.From(_glsShipmentStore.GetAllSnapshot());
                var box = GlsSubiektSync.MatchPreparingBox(_document, result.Items, cacheIndex);
                var pickups = new List<GlsPickupItem>();

                var przesylka = _przesylka ?? _document.Przesylka;
                if (box.Count == 0
                    && przesylka is { HasId: true } liveBox
                    && result.Items.Any(item => item.Id == liveBox.Id))
                {
                    box = result.Items.Where(item => item.Id == liveBox.Id).ToList();
                }

                if (przesylka is { IsEmpty: false }
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
                    var sync = await GlsSubiektSync.ApplyToLocalAsync(
                        _glsShipmentStore,
                        _document,
                        liveIds,
                        _document.GlsPreparingBoxId,
                        _document.GlsPreparingBoxParcelNumber,
                        matchedPickupNrListu: null);
                    await ApplyLocalSyncResultAndPropagateAsync(sync);
                }
                catch (Exception ex)
                {
                    Warning($"Nie udało się zsynchronizować cache GLS: {ex.Message}", "SubiektDocumentDetails");
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
                    "Brak loginu lub hasła GLS.\n\nUzupełnij je w menu Plik → Ustawienia → zakładka Ustawienia GLS.",
                    "Brak konfiguracji GLS",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            await EnsurePrzesylkaOwnerAsync(forWrite: true);

            try
            {
                await _shippingLoadTask.ConfigureAwait(true);
            }
            catch
            {
                // błąd adresu już obsłużony w LoadShippingAddressAsync
            }

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

                var savedCache = await SaveShipmentToLocalAsync(result.ConsignmentId!.Value, isEdit);
                await RefreshGlsAndSyncAsync();
                var message = isEdit
                    ? $"Przesyłka GLS ({result.EnvironmentName}) została zapisana w przygotowalni.\n\nIdentyfikator: {result.ConsignmentId}"
                    : $"Przesyłka GLS ({result.EnvironmentName}) została dodana do przygotowalni.\n\nIdentyfikator: {result.ConsignmentId}";

                if (!string.IsNullOrWhiteSpace(result.WarningMessage))
                {
                    message += "\n\n" + result.WarningMessage;
                }

                if (!string.IsNullOrWhiteSpace(savedCache.Warning))
                {
                    message += "\n\n" + savedCache.Warning;
                }

                Info(message, "SubiektDocumentDetails");
                MessageBox.Show(
                    message,
                    isEdit ? "Przygotowalnia GLS — zapisano" : "Przygotowalnia GLS — dodano",
                    MessageBoxButton.OK,
                    string.IsNullOrWhiteSpace(result.WarningMessage) && string.IsNullOrWhiteSpace(savedCache.Warning)
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
        /// w lokalnym cache i od razu na karcie przygotowalni (bez czekania na pełne Odśwież).
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
                ?? _document.GlsShipment?.NrNad
                ?? _document.GlsPickupParcelNumber);

            await EnsurePrzesylkaOwnerAsync(forWrite: true);
            try
            {
                var sync = await GlsSubiektSync.ApplyToLocalAsync(
                    _glsShipmentStore,
                    _document,
                    new HashSet<int> { consignmentId.Value },
                    consignmentId,
                    nr,
                    matchedPickupNrListu: null);
                await ApplyLocalSyncResultAndPropagateAsync(sync);
            }
            catch (Exception ex)
            {
                Warning(
                    $"Etykieta GLS ma numery {nr}, ale nie zapisano lokalnego cache GLS: {ex.Message}",
                    "SubiektDocumentDetails");
                ApplyShipmentToPrzesylka(new GlsShipmentRecord
                {
                    DokId = _document.DokId,
                    DokTyp = _document.DokTyp,
                    NrPelny = _document.NrPelny ?? "",
                    Carrier = "GLS",
                    BoxId = consignmentId.Value,
                    NrPrzyg = nr,
                    NrNad = existingNadane
                });
            }

            var box = new List<GlsPreparingBoxItem>
            {
                new()
                {
                    Id = consignmentId.Value,
                    References = _document.NrPelny ?? "",
                    ParcelNumber = nr ?? ""
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
                        References = _document.NrPelny ?? "",
                        ParcelNumber = existingNadane ?? ""
                    }
                };
                _document.GlsPickupParcelNumber = existingNadane ?? "";
            }

            WaybillStatusText = BuildGlsStatusText(PreparingBoxItems, PickupItems, pickupErrorMessage: null);
            GlsListsLoaded = true;
            UpdateWaybillButton();
            Info(
                $"Po etykiecie GLS zapisano nr listu przygotowalni: {nr} (id={consignmentId}).",
                "SubiektDocumentDetails");
        }

        private async System.Threading.Tasks.Task<(SubiektPrzesylka? Przesylka, string? Warning)> SaveShipmentToLocalAsync(
            int consignmentId,
            bool isEdit)
        {
            var nrPrzyg = isEdit
                ? SubiektPrzesylka.NormalizeNrListu(
                    !string.IsNullOrWhiteSpace(_przesylka?.NrListuPrzygotowalnia)
                        ? _przesylka!.NrListuPrzygotowalnia
                        : _document.GlsShipment?.NrPrzyg)
                : "";

            await EnsurePrzesylkaOwnerAsync(forWrite: true);

            try
            {
                var sync = await GlsSubiektSync.ApplyToLocalAsync(
                    _glsShipmentStore,
                    _document,
                    new HashSet<int> { consignmentId },
                    consignmentId,
                    nrPrzyg,
                    matchedPickupNrListu: null);
                await ApplyLocalSyncResultAndPropagateAsync(sync);
                return (_przesylka, null);
            }
            catch (Exception ex)
            {
                var warning =
                    $"Przesyłka GLS w przygotowalni ma id={consignmentId}, ale nie udało się zapisać lokalnego cache GLS: {ex.Message}";
                Warning(warning, "SubiektDocumentDetails");
                var fallback = new GlsShipmentRecord
                {
                    DokId = _document.DokId,
                    DokTyp = _document.DokTyp,
                    NrPelny = _document.NrPelny ?? "",
                    Carrier = "GLS",
                    BoxId = consignmentId,
                    NrPrzyg = nrPrzyg,
                    NrNad = SubiektPrzesylka.NormalizeNrListu(
                        _przesylka?.NrListuNadane
                        ?? _document.GlsShipment?.NrNad)
                };
                ApplyShipmentToPrzesylka(fallback);
                return (_przesylka, warning);
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
                        $"Nie udało się usunąć przesyłki z przygotowalni GLS ({deleted.EnvironmentName}).\n\nLokalny cache nie został zmieniony.\n\n{error}",
                        "Błąd GLS",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                string? cacheWarning = null;
                var refreshed = await RefreshGlsAndSyncAsync();
                if (!refreshed)
                {
                    try
                    {
                        var sync = await GlsSubiektSync.ApplyToLocalAsync(
                            _glsShipmentStore,
                            _document,
                            Array.Empty<int>(),
                            matchedPreparingBoxId: null,
                            matchedPreparingNrListu: null,
                            matchedPickupNrListu: null);
                        await ApplyLocalSyncResultAndPropagateAsync(sync);
                        _glsWaybillExists = false;
                    }
                    catch (Exception ex)
                    {
                        cacheWarning =
                            $"List usunięto z GLS, ale nie udało się zaktualizować lokalnego cache: {ex.Message}";
                        Warning(cacheWarning, "SubiektDocumentDetails");
                    }
                }

                string message;
                if (!string.IsNullOrWhiteSpace(cacheWarning))
                {
                    message = $"Przesyłkę usunięto z przygotowalni GLS ({deleted.EnvironmentName}).\n\n{cacheWarning}";
                }
                else if (deleted.NotFound)
                {
                    message = "Przesyłki nie było już w przygotowalni GLS. Lokalny cache zsynchronizowano z GLS.";
                }
                else
                {
                    message = $"Przesyłka GLS ({deleted.EnvironmentName}) została usunięta z przygotowalni. Lokalny cache zsynchronizowano z GLS.";
                }

                Info(message, "SubiektDocumentDetails");
                MessageBox.Show(
                    message,
                    "Usuń z przygotowalni",
                    MessageBoxButton.OK,
                    string.IsNullOrWhiteSpace(cacheWarning) ? MessageBoxImage.Information : MessageBoxImage.Warning);
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

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private async void EditPayerButton_Click(object sender, RoutedEventArgs e)
        {
            var khId = _document.PlatnikId > 0 ? _document.PlatnikId : _document.KontrahentId;
            if (khId <= 0)
            {
                MessageBox.Show(
                    "Brak identyfikatora kontrahenta na dokumencie.",
                    "Kontrahent",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                // Wyswietl() blokuje do zamknięcia okna Subiekta (musi być na wątku UI / STA)
                new SubiektService().OtworzKartotekeKontrahenta(khId);

                Mouse.OverrideCursor = Cursors.Wait;
                var refreshed = await _subiektApiService.GetKontrahentDetailsAsync(khId).ConfigureAwait(true);
                if (refreshed != null)
                    ApplyKontrahentDetailsToCard(refreshed);
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektDocumentDetails", "Błąd otwierania kartoteki kontrahenta w Subiekcie");
                MessageBox.Show(
                    $"Nie udało się otworzyć kartoteki kontrahenta w Subiekcie:\n\n{ex.Message}",
                    "Błąd",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void ApplyKontrahentDetailsToCard(KontrahentDetails d)
        {
            var pelna = CollapseLines(d.NazwaPelna);
            var krotka = CollapseLines(d.Nazwa);
            if (string.IsNullOrWhiteSpace(pelna))
                pelna = krotka;
            if (!string.IsNullOrWhiteSpace(krotka)
                && string.Equals(krotka, pelna, StringComparison.OrdinalIgnoreCase))
            {
                krotka = "";
            }

            PayerName = pelna;
            PayerNameShort = krotka;
            PayerNip = d.Nip?.Trim() ?? "";
            PayerEmail = d.Email?.Trim() ?? "";
            PayerPhone = d.Telefon?.Trim() ?? "";
            PayerAddress = FormatKontrahentCardAddress(d);

            _document.PlatnikNazwa = pelna;
            _document.PlatnikNazwaKrotka = krotka;
            _document.PlatnikNip = PayerNip;
            _document.PlatnikEmail = PayerEmail;
            _document.PlatnikTelefon = PayerPhone;
            _document.PlatnikAdres = PayerAddress;

            if (d.HasAdresDostawy)
            {
                var shipName = !string.IsNullOrWhiteSpace(d.DostawaNazwaPelna)
                    ? d.DostawaNazwaPelna.Trim()
                    : (d.DostawaNazwa?.Trim() ?? "");
                var shipParts = new List<string>();
                var street = !string.IsNullOrWhiteSpace(d.DostawaAdres)
                    ? d.DostawaAdres.Trim()
                    : (d.DostawaUlica?.Trim() ?? "");
                if (!string.IsNullOrWhiteSpace(street))
                    shipParts.Add(street);
                var cityLine = $"{d.DostawaKod} {d.DostawaMiejscowosc}".Trim();
                if (!string.IsNullOrWhiteSpace(cityLine))
                    shipParts.Add(cityLine);

                HasSubiektShippingAddress = true;
                SubiektShippingName = shipName;
                SubiektShippingAddress = string.Join("\n", shipParts);
                SubiektShippingPhone = d.DostawaTelefon?.Trim() ?? "";

                _document.HasAdresDostawy = true;
                _document.AdresDostawyNazwa = shipName;
                _document.AdresDostawyAdres = string.Join(", ", shipParts);
                _document.AdresDostawyUlica = street;
                _document.AdresDostawyKodPocztowy = d.DostawaKod?.Trim() ?? "";
                _document.AdresDostawyMiejscowosc = d.DostawaMiejscowosc?.Trim() ?? "";
                _document.AdresDostawyTelefon = SubiektShippingPhone;
            }
            else
            {
                HasSubiektShippingAddress = false;
                SubiektShippingName = "";
                SubiektShippingAddress = "";
                SubiektShippingPhone = "";
                _document.HasAdresDostawy = false;
                _document.AdresDostawyNazwa = "";
                _document.AdresDostawyAdres = "";
                _document.AdresDostawyUlica = "";
                _document.AdresDostawyKodPocztowy = "";
                _document.AdresDostawyMiejscowosc = "";
                _document.AdresDostawyTelefon = "";
            }
        }

        private static string FormatKontrahentCardAddress(KontrahentDetails d)
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(d.Adres))
                parts.Add(d.Adres.Trim());
            else if (!string.IsNullOrWhiteSpace(d.Ulica))
            {
                var street = d.Ulica.Trim();
                var nr = $"{d.NrDomu} {d.NrLokalu}".Trim();
                parts.Add(string.IsNullOrWhiteSpace(nr) ? street : $"{street} {nr}");
            }

            var city = $"{d.Kod} {d.Miejscowosc}".Trim();
            if (!string.IsNullOrWhiteSpace(city))
                parts.Add(city);
            return string.Join(", ", parts);
        }

        private static string CollapseLines(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";
            return string.Join(
                " ",
                value.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(p => p.Trim())
                    .Where(p => p.Length > 0));
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
