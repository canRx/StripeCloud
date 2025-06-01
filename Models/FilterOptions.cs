using System;
using System.ComponentModel;

namespace StripeCloud.Models
{
    public class FilterOptions : INotifyPropertyChanged
    {
        private string _searchText = string.Empty;
        private string? _selectedCustomer;
        private int? _selectedMonth;
        private int? _selectedYear;
        private ComparisonStatus? _selectedStatus;
        private decimal? _minAmount;
        private decimal? _maxAmount;
        private DateTime? _startDate;
        private DateTime? _endDate;

        // Freitext-Suche
        public string SearchText
        {
            get => _searchText;
            set
            {
                _searchText = value;
                OnPropertyChanged(nameof(SearchText));
            }
        }

        // Kunden-Filter
        public string? SelectedCustomer
        {
            get => _selectedCustomer;
            set
            {
                _selectedCustomer = value;
                OnPropertyChanged(nameof(SelectedCustomer));
            }
        }

        // Datum-Filter
        public int? SelectedMonth
        {
            get => _selectedMonth;
            set
            {
                _selectedMonth = value;
                OnPropertyChanged(nameof(SelectedMonth));
                UpdateDateRange();
            }
        }

        public int? SelectedYear
        {
            get => _selectedYear;
            set
            {
                _selectedYear = value;
                OnPropertyChanged(nameof(SelectedYear));
                UpdateDateRange();
            }
        }

        public DateTime? StartDate
        {
            get => _startDate;
            set
            {
                _startDate = value;
                OnPropertyChanged(nameof(StartDate));
            }
        }

        public DateTime? EndDate
        {
            get => _endDate;
            set
            {
                _endDate = value;
                OnPropertyChanged(nameof(EndDate));
            }
        }

        // Status-Filter
        public ComparisonStatus? SelectedStatus
        {
            get => _selectedStatus;
            set
            {
                _selectedStatus = value;
                OnPropertyChanged(nameof(SelectedStatus));
            }
        }

        // Betrag-Filter
        public decimal? MinAmount
        {
            get => _minAmount;
            set
            {
                _minAmount = value;
                OnPropertyChanged(nameof(MinAmount));
            }
        }

        public decimal? MaxAmount
        {
            get => _maxAmount;
            set
            {
                _maxAmount = value;
                OnPropertyChanged(nameof(MaxAmount));
            }
        }

        // Computed Properties
        public bool HasAnyFilter =>
            !string.IsNullOrWhiteSpace(SearchText) ||
            !string.IsNullOrWhiteSpace(SelectedCustomer) ||
            SelectedMonth.HasValue ||
            SelectedYear.HasValue ||
            SelectedStatus.HasValue ||
            MinAmount.HasValue ||
            MaxAmount.HasValue ||
            StartDate.HasValue ||
            EndDate.HasValue;

        public string FilterSummary
        {
            get
            {
                if (!HasAnyFilter) return "Keine Filter aktiv";

                var filters = new List<string>();

                if (!string.IsNullOrWhiteSpace(SearchText))
                    filters.Add($"Suche: '{SearchText}'");

                if (!string.IsNullOrWhiteSpace(SelectedCustomer))
                    filters.Add($"Kunde: {SelectedCustomer}");

                if (SelectedMonth.HasValue && SelectedYear.HasValue)
                    filters.Add($"Datum: {SelectedMonth:D2}/{SelectedYear}");
                else if (SelectedYear.HasValue)
                    filters.Add($"Jahr: {SelectedYear}");
                else if (SelectedMonth.HasValue)
                    filters.Add($"Monat: {SelectedMonth:D2}");

                if (StartDate.HasValue || EndDate.HasValue)
                {
                    if (StartDate.HasValue && EndDate.HasValue)
                        filters.Add($"Zeitraum: {StartDate:dd.MM.yyyy} - {EndDate:dd.MM.yyyy}");
                    else if (StartDate.HasValue)
                        filters.Add($"Ab: {StartDate:dd.MM.yyyy}");
                    else if (EndDate.HasValue)
                        filters.Add($"Bis: {EndDate:dd.MM.yyyy}");
                }

                if (SelectedStatus.HasValue)
                    filters.Add($"Status: {GetStatusDisplayName(SelectedStatus.Value)}");

                if (MinAmount.HasValue || MaxAmount.HasValue)
                {
                    if (MinAmount.HasValue && MaxAmount.HasValue)
                        filters.Add($"Betrag: {MinAmount:C} - {MaxAmount:C}");
                    else if (MinAmount.HasValue)
                        filters.Add($"Min. Betrag: {MinAmount:C}");
                    else if (MaxAmount.HasValue)
                        filters.Add($"Max. Betrag: {MaxAmount:C}");
                }

                return string.Join(", ", filters);
            }
        }

        // Helper Methods
        private void UpdateDateRange()
        {
            if (SelectedYear.HasValue && SelectedMonth.HasValue)
            {
                StartDate = new DateTime(SelectedYear.Value, SelectedMonth.Value, 1);
                EndDate = StartDate.Value.AddMonths(1).AddDays(-1);
            }
            else if (SelectedYear.HasValue)
            {
                StartDate = new DateTime(SelectedYear.Value, 1, 1);
                EndDate = new DateTime(SelectedYear.Value, 12, 31);
            }
            else
            {
                StartDate = null;
                EndDate = null;
            }
        }

        private static string GetStatusDisplayName(ComparisonStatus status)
        {
            return status switch
            {
                ComparisonStatus.Match => "Übereinstimmung",
                ComparisonStatus.OnlyStripe => "Nur Stripe",
                ComparisonStatus.OnlyChargecloud => "Nur Chargecloud",
                ComparisonStatus.AmountMismatch => "Betragsabweichung",
                _ => status.ToString()
            };
        }

        // Filter zurücksetzen
        public void ClearAllFilters()
        {
            SearchText = string.Empty;
            SelectedCustomer = null;
            SelectedMonth = null;
            SelectedYear = null;
            SelectedStatus = null;
            MinAmount = null;
            MaxAmount = null;
            StartDate = null;
            EndDate = null;
        }

        // Prüfen ob eine Transaktion den Filtern entspricht
        public bool MatchesFilter(TransactionComparison transaction)
        {
            // Freitext-Suche
            if (!string.IsNullOrWhiteSpace(SearchText))
            {
                var searchLower = SearchText.ToLowerInvariant();
                if (!transaction.CustomerEmail.ToLowerInvariant().Contains(searchLower) &&
                    !transaction.DisplayName.ToLowerInvariant().Contains(searchLower) &&
                    !transaction.StatusText.ToLowerInvariant().Contains(searchLower))
                {
                    return false;
                }
            }

            // Kunden-Filter
            if (!string.IsNullOrWhiteSpace(SelectedCustomer) &&
                !transaction.CustomerEmail.Equals(SelectedCustomer, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            // Datum-Filter
            if (StartDate.HasValue && transaction.TransactionDate.Date < StartDate.Value.Date)
                return false;

            if (EndDate.HasValue && transaction.TransactionDate.Date > EndDate.Value.Date)
                return false;

            // Status-Filter
            if (SelectedStatus.HasValue && transaction.Status != SelectedStatus.Value)
                return false;

            // Betrag-Filter
            if (MinAmount.HasValue && Math.Abs(transaction.Amount) < MinAmount.Value)
                return false;

            if (MaxAmount.HasValue && Math.Abs(transaction.Amount) > MaxAmount.Value)
                return false;

            return true;
        }

        // INotifyPropertyChanged Implementation
        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}