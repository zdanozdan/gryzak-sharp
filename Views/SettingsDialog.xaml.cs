using System.Windows;
using Gryzak.Services;

namespace Gryzak.Views
{
    public enum SettingsTab
    {
        Shop,
        Subiekt,
        Gls,
        Ai
    }

    public partial class SettingsDialog : Window
    {
        private readonly ShopSettingsPanel _shopPanel;
        private readonly SubiektSettingsPanel _subiektPanel;
        private readonly GlsSettingsPanel _glsPanel;
        private readonly AiSettingsPanel _aiPanel;

        public SettingsDialog(ConfigService configService, SettingsTab initialTab = SettingsTab.Shop)
        {
            InitializeComponent();

            _shopPanel = new ShopSettingsPanel(configService);
            _subiektPanel = new SubiektSettingsPanel(configService);
            _glsPanel = new GlsSettingsPanel(configService);
            _aiPanel = new AiSettingsPanel(configService);

            ShopContentHost.Content = _shopPanel;
            SubiektContentHost.Content = _subiektPanel;
            GlsContentHost.Content = _glsPanel;
            AiContentHost.Content = _aiPanel;

            SelectTab(initialTab);
        }

        private void SelectTab(SettingsTab tab)
        {
            SettingsTabControl.SelectedItem = tab switch
            {
                SettingsTab.Subiekt => SubiektTab,
                SettingsTab.Gls => GlsTab,
                SettingsTab.Ai => AiTab,
                _ => ShopTab
            };
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_shopPanel.TrySave())
            {
                SettingsTabControl.SelectedItem = ShopTab;
                return;
            }

            if (!_subiektPanel.TrySave())
            {
                SettingsTabControl.SelectedItem = SubiektTab;
                return;
            }

            if (!_glsPanel.TrySave())
            {
                SettingsTabControl.SelectedItem = GlsTab;
                return;
            }

            if (!_aiPanel.TrySave())
            {
                SettingsTabControl.SelectedItem = AiTab;
                return;
            }

            MessageBox.Show(
                "Ustawienia zostały zapisane pomyślnie.",
                "Sukces",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
