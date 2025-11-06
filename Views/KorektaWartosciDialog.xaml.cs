using System.Windows;

namespace Gryzak.Views
{
    public partial class KorektaWartosciDialog : Window
    {
        public bool CzyKorygowac { get; private set; } = false;
        public bool CzyAnulowac { get; private set; } = false;

        public KorektaWartosciDialog(string komunikat, bool hasCoupon = false, string? couponTitle = null, double? couponAmount = null, string currency = "PLN")
        {
            InitializeComponent();
            MessageTextBlock.Text = komunikat;
            
            // Wyświetl informację o kuponie, jeśli jest dostępny
            if (hasCoupon && couponAmount.HasValue && couponAmount.Value > 0.01)
            {
                CouponInfoPanel.Visibility = Visibility.Visible;
                string couponText = "";
                if (!string.IsNullOrWhiteSpace(couponTitle))
                {
                    couponText = $"Kupon rabatowy: {couponTitle} (-{couponAmount.Value:F2} {currency})";
                }
                else
                {
                    couponText = $"Kupon rabatowy: -{couponAmount.Value:F2} {currency}";
                }
                CouponInfoText.Text = couponText;
            }
            else
            {
                CouponInfoPanel.Visibility = Visibility.Collapsed;
            }
        }

        private void KorygujButton_Click(object sender, RoutedEventArgs e)
        {
            CzyKorygowac = true;
            CzyAnulowac = false;
            DialogResult = true;
            Close();
        }

        private void PozostawButton_Click(object sender, RoutedEventArgs e)
        {
            CzyKorygowac = false;
            CzyAnulowac = false;
            DialogResult = false;
            Close();
        }
        
        private void AnulujButton_Click(object sender, RoutedEventArgs e)
        {
            CzyKorygowac = false;
            CzyAnulowac = true;
            DialogResult = false;
            Close();
        }
    }
}

