using System.Windows;

namespace Gryzak.Views
{
    public partial class KorektaWartosciDialog : Window
    {
        public bool CzyKorygowac { get; private set; } = false;

        public KorektaWartosciDialog(string komunikat)
        {
            InitializeComponent();
            MessageTextBlock.Text = komunikat;
        }

        private void KorygujButton_Click(object sender, RoutedEventArgs e)
        {
            CzyKorygowac = true;
            DialogResult = true;
            Close();
        }

        private void PozostawButton_Click(object sender, RoutedEventArgs e)
        {
            CzyKorygowac = false;
            DialogResult = false;
            Close();
        }
    }
}

