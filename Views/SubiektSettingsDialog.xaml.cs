using System.Windows;
using Gryzak.Models;
using Gryzak.Services;

namespace Gryzak.Views
{
    public partial class SubiektSettingsDialog : Window
    {
        private readonly ConfigService _configService;
        private SubiektConfig _currentConfig;

        public SubiektSettingsDialog(ConfigService configService)
        {
            InitializeComponent();
            _configService = configService;
            _currentConfig = _configService.LoadSubiektConfig();
            LoadConfig();
        }

        private void LoadConfig()
        {
            UserTextBox.Text = _currentConfig.User;
            PasswordBox.Password = _currentConfig.Password;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            _currentConfig.User = UserTextBox.Text.Trim();
            _currentConfig.Password = PasswordBox.Password;

            try
            {
                _configService.SaveSubiektConfig(_currentConfig);
                MessageBox.Show("Ustawienia Subiekt GT zostały zapisane pomyślnie.", "Sukces", MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = true;
                Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Nie udało się zapisać ustawień:\n\n{ex.Message}", "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}

