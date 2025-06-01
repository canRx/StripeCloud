using System.Windows;

namespace StripeCloud.Views
{
    public partial class TransactionDetailWindow : Window
    {
        public TransactionDetailWindow()
        {
            InitializeComponent();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}