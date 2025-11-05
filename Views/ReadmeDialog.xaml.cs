using System;
using System.IO;
using System.Windows;

namespace Gryzak.Views
{
    public partial class ReadmeDialog : Window
    {
        public ReadmeDialog()
        {
            InitializeComponent();
            Loaded += ReadmeDialog_Loaded;
        }

        private void ReadmeDialog_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Wczytaj README.md
                string readmePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "README.md");
                if (File.Exists(readmePath))
                {
                    string readmeContent = File.ReadAllText(readmePath);
                    ReadmeTextBlock.Text = readmeContent;
                }
                else
                {
                    // Jeśli plik nie istnieje w katalogu aplikacji, spróbuj znaleźć go w katalogu projektu
                    string projectReadmePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "README.md");
                    if (File.Exists(projectReadmePath))
                    {
                        string readmeContent = File.ReadAllText(projectReadmePath);
                        ReadmeTextBlock.Text = readmeContent;
                    }
                    else
                    {
                        ReadmeTextBlock.Text = "Nie można znaleźć pliku README.md";
                    }
                }

                // Wczytaj INSTALLER.md
                string installerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "INSTALLER.md");
                if (File.Exists(installerPath))
                {
                    string installerContent = File.ReadAllText(installerPath);
                    InstallerTextBlock.Text = installerContent;
                }
                else
                {
                    // Jeśli plik nie istnieje w katalogu aplikacji, spróbuj znaleźć go w katalogu projektu
                    string projectInstallerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "INSTALLER.md");
                    if (File.Exists(projectInstallerPath))
                    {
                        string installerContent = File.ReadAllText(projectInstallerPath);
                        InstallerTextBlock.Text = installerContent;
                    }
                    else
                    {
                        InstallerTextBlock.Text = "Nie można znaleźć pliku INSTALLER.md";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Błąd podczas wczytywania dokumentacji:\n\n{ex.Message}",
                    "Błąd",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

