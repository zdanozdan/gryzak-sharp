using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Gryzak.Models;
using Gryzak.Services;
using PDFtoImage;
using SkiaSharp;
using static Gryzak.Services.Logger;

namespace Gryzak.Views
{
    public partial class GlsLabelDialog : Window
    {
        private readonly GlsConfig _glsConfig;
        private readonly ConfigService _configService;
        private readonly int _consignmentId;
        private byte[]? _pdfBytes;
        private byte[]? _zplBytes;
        private bool _isBusy;
        private double _zoomFactor = 1.0; // 1.0 = dopasowanie do okna
        private const double ZoomMin = 0.25;
        private const double ZoomMax = 6.0;
        private const double ZoomStep = 0.25;

        /// <summary>True, gdy GLS zwróciło PDF i/lub ZPL — wtedy numery paczek są już nadane.</summary>
        public bool LabelsLoaded => _pdfBytes is { Length: > 0 } || _zplBytes is { Length: > 0 };

        public GlsLabelDialog(int consignmentId, GlsConfig glsConfig, ConfigService? configService = null)
        {
            InitializeComponent();
            _consignmentId = consignmentId;
            _glsConfig = glsConfig ?? new GlsConfig();
            _configService = configService ?? new ConfigService();

            SubtitleText.Text =
                $"Id w przygotowalni: {_consignmentId}  •  Środowisko: {_glsConfig.GetEnvironmentName()}"
                + "\nPodgląd: PDF  •  Druk na Zebrę: ZPL (RAW)";

            OrientationVerticalRadio.IsChecked = !_glsConfig.LabelPreviewLandscape;
            OrientationHorizontalRadio.IsChecked = _glsConfig.LabelPreviewLandscape;
            ApplyOrientation(_glsConfig.LabelPreviewLandscape, resetZoom: false);
            UpdateZoomUi();

            LoadPrinters();
            Loaded += async (_, _) => await LoadLabelsAsync();
        }

        private void Orientation_Changed(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded)
            {
                return;
            }

            var landscape = OrientationHorizontalRadio.IsChecked == true;
            ApplyOrientation(landscape, resetZoom: true);
            PersistOrientation(landscape);
        }

        private void ApplyOrientation(bool landscape, bool resetZoom)
        {
            // PDF z GLS (roll 160×100) jest w układzie poziomym — „pionowa” = obrót 90°.
            PreviewRotate.Angle = landscape ? 0 : 90;
            if (resetZoom)
            {
                _zoomFactor = 1.0;
            }

            _ = Dispatcher.BeginInvoke(new Action(ApplyZoom), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        private void ZoomInButton_Click(object sender, RoutedEventArgs e) => ChangeZoom(_zoomFactor + ZoomStep);

        private void ZoomOutButton_Click(object sender, RoutedEventArgs e) => ChangeZoom(_zoomFactor - ZoomStep);

        private void ZoomFitButton_Click(object sender, RoutedEventArgs e) => ChangeZoom(1.0);

        private void PreviewScroll_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.Control || PreviewImage.Source == null)
            {
                return;
            }

            e.Handled = true;
            ChangeZoom(_zoomFactor + (e.Delta > 0 ? ZoomStep : -ZoomStep));
        }

        private void PreviewScroll_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (PreviewImage.Source != null)
            {
                ApplyZoom();
            }
        }

        private void ChangeZoom(double factor)
        {
            _zoomFactor = Math.Clamp(factor, ZoomMin, ZoomMax);
            ApplyZoom();
        }

        private void ApplyZoom()
        {
            if (PreviewImage.Source is not BitmapSource bmp)
            {
                PreviewScale.ScaleX = PreviewScale.ScaleY = 1;
                UpdateZoomUi();
                return;
            }

            var landscape = PreviewRotate.Angle == 0;
            var imgW = landscape ? bmp.PixelWidth : bmp.PixelHeight;
            var imgH = landscape ? bmp.PixelHeight : bmp.PixelWidth;
            if (imgW <= 0 || imgH <= 0)
            {
                return;
            }

            var availW = Math.Max(1, PreviewScroll.ViewportWidth - 8);
            var availH = Math.Max(1, PreviewScroll.ViewportHeight - 8);
            if (availW < 2 || availH < 2)
            {
                availW = Math.Max(1, PreviewScroll.ActualWidth - 8);
                availH = Math.Max(1, PreviewScroll.ActualHeight - 8);
            }

            var fit = Math.Min(availW / imgW, availH / imgH);
            if (double.IsNaN(fit) || double.IsInfinity(fit) || fit <= 0)
            {
                fit = 1;
            }

            // Przy zoomie 100% (fit) nie powiększamy ponad naturalny rozmiar, jeśli etykieta jest mniejsza niż okno.
            // Ale user chciał fit-to-window z powiększeniem — więc pozwalamy fit > 1.
            var scale = fit * _zoomFactor;
            PreviewScale.ScaleX = scale;
            PreviewScale.ScaleY = scale;
            UpdateZoomUi();
        }

        private void UpdateZoomUi()
        {
            ZoomLevelText.Text = $"{(_zoomFactor * 100):0}%";
            ZoomOutButton.IsEnabled = _zoomFactor > ZoomMin + 0.001;
            ZoomInButton.IsEnabled = _zoomFactor < ZoomMax - 0.001;
        }

        private void LoadPrinters()
        {
            var printers = RawPrinterHelper.GetInstalledPrinterNames();
            PrinterComboBox.ItemsSource = printers;
            var preferred = RawPrinterHelper.PickPreferredPrinter(_glsConfig.LabelPrinterName, printers);
            if (preferred != null)
            {
                PrinterComboBox.SelectedItem = preferred;
            }
            else if (printers.Count > 0)
            {
                PrinterComboBox.SelectedIndex = 0;
            }
        }

        private async System.Threading.Tasks.Task LoadLabelsAsync()
        {
            if (_isBusy)
            {
                return;
            }

            _isBusy = true;
            PrintButton.IsEnabled = false;
            PreviewStatusText.Visibility = Visibility.Visible;
            PreviewStatusText.Text = "Pobieranie etykiety z GLS...";
            PreviewImage.Source = null;
            StatusText.Text = "";
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                using var gls = new GlsService(_glsConfig);

                var pdfTask = gls.GetLabelsAsync(_consignmentId, GlsLabelResult.PreviewPdfMode);
                var zplTask = gls.GetLabelsAsync(_consignmentId, GlsLabelResult.ZebraZplMode);
                await System.Threading.Tasks.Task.WhenAll(pdfTask, zplTask);

                var pdf = await pdfTask;
                var zpl = await zplTask;

                if ((!pdf.Success || pdf.LabelBytes is not { Length: > 0 })
                    && !string.Equals(GlsLabelResult.PreviewPdfMode, GlsLabelResult.PdfMode, StringComparison.Ordinal))
                {
                    // Fallback: A4 + crop białych marginesów
                    var a4 = await gls.GetLabelsAsync(_consignmentId, GlsLabelResult.PdfMode);
                    if (a4.Success && a4.LabelBytes is { Length: > 0 })
                    {
                        pdf = a4;
                    }
                }

                if (pdf.Success && pdf.LabelBytes is { Length: > 0 })
                {
                    _pdfBytes = pdf.LabelBytes;
                    try
                    {
                        PreviewImage.Source = RenderPdfPreview(_pdfBytes);
                        PreviewStatusText.Visibility = Visibility.Collapsed;
                        _zoomFactor = 1.0;
                        _ = Dispatcher.BeginInvoke(new Action(ApplyZoom), System.Windows.Threading.DispatcherPriority.Loaded);
                    }
                    catch (Exception ex)
                    {
                        Warning($"Nie udało się wyrenderować podglądu PDF: {ex.Message}", "GlsLabelDialog");
                        PreviewStatusText.Text = "Pobrano PDF, ale podgląd nie udał się. Możesz nadal drukować ZPL.";
                    }
                }
                else
                {
                    PreviewStatusText.Text = pdf.ErrorMessage ?? "Nie pobrano PDF do podglądu.";
                }

                if (zpl.Success && zpl.LabelBytes is { Length: > 0 })
                {
                    _zplBytes = zpl.LabelBytes;
                }
                else
                {
                    StatusText.Text = zpl.ErrorMessage ?? "Nie pobrano ZPL do druku.";
                }

                if (_zplBytes != null)
                {
                    PrintButton.IsEnabled = true;
                    if (string.IsNullOrWhiteSpace(StatusText.Text))
                    {
                        StatusText.Text = $"Gotowe do druku (ZPL {_zplBytes.Length} B).";
                    }
                    else if (_pdfBytes != null)
                    {
                        StatusText.Text = $"Podgląd OK. ZPL: {zpl.ErrorMessage}";
                    }
                }
                else if (_pdfBytes != null)
                {
                    StatusText.Text = "Jest podgląd PDF, ale brak ZPL — druk na Zebrę niedostępny.";
                }

                Info(
                    $"Etykieta GLS id={_consignmentId}: pdf={_pdfBytes?.Length ?? 0} B, zpl={_zplBytes?.Length ?? 0} B",
                    "GlsLabelDialog");
            }
            catch (Exception ex)
            {
                Error(ex, "GlsLabelDialog", "Błąd pobierania etykiety");
                PreviewStatusText.Text = $"Błąd pobierania: {ex.Message}";
                StatusText.Text = ex.Message;
            }
            finally
            {
                Mouse.OverrideCursor = null;
                _isBusy = false;
            }
        }

        private static BitmapSource RenderPdfPreview(byte[] pdfBytes)
        {
            using var page = Conversion.ToImage(pdfBytes, options: new(Dpi: 200));
            using var cropped = CropToContent(page) ?? page.Copy();
            return ToBitmapSource(cropped);
        }

        /// <summary>
        /// Obcina białe marginesy kartki — zostaje sama etykieta do fit-to-window.
        /// </summary>
        private static SKBitmap? CropToContent(SKBitmap source, byte whiteThreshold = 245, int padding = 4)
        {
            var width = source.Width;
            var height = source.Height;
            if (width <= 0 || height <= 0)
            {
                return null;
            }

            var minX = width;
            var minY = height;
            var maxX = -1;
            var maxY = -1;

            for (var y = 0; y < height; y++)
            {
                for (var x = 0; x < width; x++)
                {
                    var c = source.GetPixel(x, y);
                    if (c.Alpha < 16)
                    {
                        continue;
                    }

                    if (c.Red >= whiteThreshold && c.Green >= whiteThreshold && c.Blue >= whiteThreshold)
                    {
                        continue;
                    }

                    if (x < minX) minX = x;
                    if (y < minY) minY = y;
                    if (x > maxX) maxX = x;
                    if (y > maxY) maxY = y;
                }
            }

            if (maxX < minX || maxY < minY)
            {
                return null;
            }

            minX = Math.Max(0, minX - padding);
            minY = Math.Max(0, minY - padding);
            maxX = Math.Min(width - 1, maxX + padding);
            maxY = Math.Min(height - 1, maxY + padding);

            var cropWidth = maxX - minX + 1;
            var cropHeight = maxY - minY + 1;

            // Jeśli „treść” zajmuje prawie całą stronę — nie ma sensu cropować.
            if (cropWidth > width * 0.95 && cropHeight > height * 0.95)
            {
                return null;
            }

            var cropped = new SKBitmap(cropWidth, cropHeight);
            using var canvas = new SKCanvas(cropped);
            canvas.DrawBitmap(
                source,
                new SKRect(minX, minY, maxX + 1, maxY + 1),
                new SKRect(0, 0, cropWidth, cropHeight));
            return cropped;
        }

        private static BitmapSource ToBitmapSource(SKBitmap bitmap)
        {
            var info = bitmap.Info;
            var pixels = bitmap.GetPixelSpan().ToArray();
            var stride = info.RowBytes;
            var source = BitmapSource.Create(
                info.Width,
                info.Height,
                96,
                96,
                PixelFormats.Bgra32,
                null,
                pixels,
                stride);
            source.Freeze();
            return source;
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isBusy)
            {
                return;
            }

            var printer = PrinterComboBox.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(printer))
            {
                MessageBox.Show("Wybierz drukarkę etykiet.", "Brak drukarki", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_zplBytes == null || _zplBytes.Length == 0)
            {
                MessageBox.Show(
                    "Brak danych ZPL do druku. Zamknij okno i spróbuj ponownie.",
                    "Brak ZPL",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            _isBusy = true;
            PrintButton.IsEnabled = false;
            PrintButton.Content = "Drukowanie...";
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                var payload = EnsureZplBytes(_zplBytes);
                if (!RawPrinterHelper.SendBytes(printer, payload, $"GLS-{_consignmentId}", out var error))
                {
                    Warning(error ?? "Błąd druku", "GlsLabelDialog");
                    MessageBox.Show(
                        $"Nie udało się wydrukować etykiety na „{printer}”.\n\n{error}",
                        "Błąd druku",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                PersistPreferredPrinter(printer);
                StatusText.Text = $"Wysłano ZPL na „{printer}”.";
                Info(StatusText.Text, "GlsLabelDialog");
                MessageBox.Show(
                    $"Etykieta została wysłana na drukarkę:\n{printer}",
                    "Wydrukowano",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Error(ex, "GlsLabelDialog", "Błąd druku etykiety");
                MessageBox.Show($"Błąd druku:\n\n{ex.Message}", "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
                PrintButton.Content = "Drukuj";
                PrintButton.IsEnabled = _zplBytes != null;
                _isBusy = false;
            }
        }

        private static byte[] EnsureZplBytes(byte[] raw)
        {
            // Jeśli API oddało tekst ZPL jako UTF-8/ASCII — wyślij jak jest.
            // BOM UTF-8 odetnij.
            if (raw.Length >= 3 && raw[0] == 0xEF && raw[1] == 0xBB && raw[2] == 0xBF)
            {
                return raw.Skip(3).ToArray();
            }

            return raw;
        }

        private void PersistPreferredPrinter(string printer)
        {
            try
            {
                if (string.Equals(_glsConfig.LabelPrinterName, printer, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                _glsConfig.LabelPrinterName = printer;
                _configService.SaveGlsConfig(_glsConfig);
            }
            catch (Exception ex)
            {
                Warning($"Nie zapisano preferowanej drukarki: {ex.Message}", "GlsLabelDialog");
            }
        }

        private void PersistOrientation(bool landscape)
        {
            try
            {
                if (_glsConfig.LabelPreviewLandscape == landscape)
                {
                    return;
                }

                _glsConfig.LabelPreviewLandscape = landscape;
                _configService.SaveGlsConfig(_glsConfig);
            }
            catch (Exception ex)
            {
                Warning($"Nie zapisano orientacji podglądu: {ex.Message}", "GlsLabelDialog");
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
