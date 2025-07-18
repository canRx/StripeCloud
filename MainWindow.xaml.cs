using StripeCloud.Models;
using StripeCloud.ViewModels;
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

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

        // Sortier-Event Handler
        private void DataGrid_Sorting(object sender, DataGridSortingEventArgs e)
        {
            if (DataContext is not MainViewModel viewModel)
                return;

            // Verhindere das Standard-Sortier-Verhalten
            e.Handled = true;

            var column = e.Column;
            var sortMemberPath = column.SortMemberPath;

            if (string.IsNullOrEmpty(sortMemberPath))
                return;

            // Bestimme die neue Sortierrichtung
            ListSortDirection direction;
            if (column.SortDirection == null || column.SortDirection == ListSortDirection.Descending)
            {
                direction = ListSortDirection.Ascending;
                column.SortDirection = ListSortDirection.Ascending;
            }
            else
            {
                direction = ListSortDirection.Descending;
                column.SortDirection = ListSortDirection.Descending;
            }

            // Lösche andere Sortier-Indikatoren
            var dataGrid = sender as DataGrid;
            if (dataGrid != null)
            {
                foreach (var col in dataGrid.Columns)
                {
                    if (col != column)
                        col.SortDirection = null;
                }
            }

            // Rufe die Sortier-Methode im ViewModel auf
            viewModel.HandleDataGridSorting(sortMemberPath, direction);
        }

        // Event-Handler für Checkbox-Auswahl
        private void TransactionCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox checkBox && checkBox.DataContext is TransactionComparison transaction)
            {
                var viewModel = DataContext as MainViewModel;
                viewModel?.ToggleTransactionSelectionCommand.Execute(transaction);
            }
        }

        private void TransactionCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox checkBox && checkBox.DataContext is TransactionComparison transaction)
            {
                var viewModel = DataContext as MainViewModel;
                viewModel?.ToggleTransactionSelectionCommand.Execute(transaction);
            }
        }
    }
}