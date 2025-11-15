using System.Windows;

namespace revit_mcp_plugin.UI
{
    public partial class ResultsWindow : Window
    {
        public string ResultText
        {
            get => ResultTextBox.Text;
            set => ResultTextBox.Text = value;
        }

        public ResultsWindow()
        {
            InitializeComponent();
        }

        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(ResultTextBox.Text))
            {
                Clipboard.SetText(ResultTextBox.Text);
                MessageBox.Show("Results copied to clipboard!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}

