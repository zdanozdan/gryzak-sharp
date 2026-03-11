using System;
using System.Windows;

namespace Gryzak.Views
{
    public partial class ProgressDialog : Window
    {
        public bool IsCancelled { get; private set; } = false;
        private bool _isStarted = false;
        private bool _isFinished = false;
        private System.Threading.Tasks.TaskCompletionSource<bool> _startTcs = new System.Threading.Tasks.TaskCompletionSource<bool>();

        public ProgressDialog()
        {
            InitializeComponent();
        }

        public async System.Threading.Tasks.Task WaitForStart()
        {
            await _startTcs.Task;
        }

        public void UpdateProgress(double value, string status, string countText = "")
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBar.Value = value;
                StatusTextBlock.Text = status;
                
                if (!string.IsNullOrEmpty(countText))
                {
                    ProgressCountTextBlock.Text = countText;
                }

                if (value >= 100)
                {
                    MarkAsFinished();
                }
            });
        }

        public void MarkAsFinished()
        {
            Dispatcher.Invoke(() =>
            {
                _isFinished = true;
                ActionButton.Content = "Gotowe";
                Title = "Zakończono";
            });
        }

        private void ActionButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_isStarted)
            {
                _isStarted = true;
                ActionButton.Content = "Przerwij";
                _startTcs.SetResult(true);
                return;
            }

            if (_isFinished || IsCancelled)
            {
                Close();
            }
            else
            {
                IsCancelled = true;
                StatusTextBlock.Text = "Przerywanie...";
                ActionButton.Content = "Gotowe";
                Title = "Przerwano";
                _isFinished = true; // Pozwól zamknąć przyciskiem
            }
        }
    }
}
