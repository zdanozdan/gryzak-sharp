using System.Windows;

namespace Gryzak.Views
{
    public partial class KorektaWartosciDialog : Window
    {
        public bool CzyKorygowac { get; private set; } = false;
        public bool CzyAnulowac { get; private set; } = false;

        public KorektaWartosciDialog(string komunikat)
        {
            InitializeComponent();
            MessageTextBlock.Text = komunikat;
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

