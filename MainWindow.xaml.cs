using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using StripeCloud.Models;
using StripeCloud.ViewModels;

namespace StripeCloud
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is DataGrid dataGrid && dataGrid.SelectedItem is TransactionComparison transaction)
            {
                var viewModel = DataContext as MainViewModel;
                viewModel?.OpenDetailWindowCommand.Execute(null);
            }
        }
    }
}