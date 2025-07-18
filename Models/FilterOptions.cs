using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace StripeCloud.Models
{
    public class FilterOptions : INotifyPropertyChanged
    {
        private string _searchText = string.Empty;
        private string? _selectedCustomer;
        private int? _selectedMonth;
        private int? _selectedYear;
        private ComparisonStatus? _selectedStatus;
        private StatusFilterOption? _selectedStatusFilter;
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

        // Status-Filter (alt)
        public ComparisonStatus? SelectedStatus
        {
            get => _selectedStatus;
            set
            {
                _selectedStatus = value;
                OnPropertyChanged(nameof(SelectedStatus));
            }
        }

        // Erweiterte Status-Filter
        public StatusFilterOption? SelectedStatusFilter
        {
            get => _selectedStatusFilter;
            set
            {
                _selectedStatusFilter = value;
                // Automatisch den alten Status setzen für Kompatibilität
                _selectedStatus = value?.Status;
                OnPropertyChanged(nameof(SelectedStatusFilter));
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
            SelectedStatusFilter != null ||
            MinAmount.HasValue ||
            MaxAmount.HasValue ||
            StartDate.HasValue ||
            EndDate.HasValue;

        // Helper Methods
        private void UpdateDateRange()
        {
            // Nur wenn BEIDE Werte gesetzt sind, setze StartDate/EndDate
            if (SelectedYear.HasValue && SelectedMonth.HasValue)
            {
                StartDate = new DateTime(SelectedYear.Value, SelectedMonth.Value, 1);
                EndDate = StartDate.Value.AddMonths(1).AddDays(-1);
            }
            else if (SelectedYear.HasValue && !SelectedMonth.HasValue)
            {
                // Nur Jahr gesetzt - ganzes Jahr
                StartDate = new DateTime(SelectedYear.Value, 1, 1);
                EndDate = new DateTime(SelectedYear.Value, 12, 31);
            }
            else
            {
                // Nur Monat oder gar nichts - keine Datumsbereich-Filterung
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
            SelectedStatusFilter = null;
            MinAmount = null;
            MaxAmount = null;
            StartDate = null;
            EndDate = null;
        }

        // Schnellfilter-Methoden
        public void SetOnlyMatches()
        {
            SelectedStatusFilter = StatusFilterOption.OnlyMatches;
        }

        public void SetOnlyProblems()
        {
            SelectedStatusFilter = StatusFilterOption.OnlyProblems;
        }

        public void SetOnlyAmountMismatch()
        {
            SelectedStatusFilter = StatusFilterOption.AmountMismatch;
        }

        public void SetManuallyConfirmed()
        {
            SelectedStatusFilter = StatusFilterOption.ManuallyConfirmed;
        }

        public void ShowAll()
        {
            SelectedStatusFilter = null;
            SelectedStatus = null;
        }

        // GEÄNDERT: Neue Filter-Logik - Monat/Jahr unabhängig
        public bool MatchesFilter(TransactionComparison transaction)
        {
            var filterResults = new List<bool>();

            // Freitext-Suche (wenn angegeben)
            if (!string.IsNullOrWhiteSpace(SearchText))
            {
                var searchLower = SearchText.ToLowerInvariant();
                var textMatches = transaction.CustomerEmail.ToLowerInvariant().Contains(searchLower) ||
                                transaction.DisplayName.ToLowerInvariant().Contains(searchLower) ||
                                transaction.StatusText.ToLowerInvariant().Contains(searchLower);
                filterResults.Add(textMatches);
            }

            // Kunden-Filter (wenn angegeben)
            if (!string.IsNullOrWhiteSpace(SelectedCustomer))
            {
                var customerMatches = transaction.CustomerEmail.Equals(SelectedCustomer, StringComparison.OrdinalIgnoreCase);
                filterResults.Add(customerMatches);
            }

            // GEÄNDERT: Monat-Filter unabhängig vom Jahr
            if (SelectedMonth.HasValue)
            {
                var monthMatches = transaction.TransactionDate.Month == SelectedMonth.Value;
                filterResults.Add(monthMatches);
            }

            // GEÄNDERT: Jahr-Filter unabhängig vom Monat
            if (SelectedYear.HasValue)
            {
                var yearMatches = transaction.TransactionDate.Year == SelectedYear.Value;
                filterResults.Add(yearMatches);
            }

            // Datum-Filter über StartDate/EndDate (nur wenn NICHT über Monat/Jahr gesetzt)
            if (StartDate.HasValue && !SelectedMonth.HasValue && !SelectedYear.HasValue)
            {
                var afterStartDate = transaction.TransactionDate.Date >= StartDate.Value.Date;
                filterResults.Add(afterStartDate);
            }

            if (EndDate.HasValue && !SelectedMonth.HasValue && !SelectedYear.HasValue)
            {
                var beforeEndDate = transaction.TransactionDate.Date <= EndDate.Value.Date;
                filterResults.Add(beforeEndDate);
            }

            // Status-Filter (erweitert) - wenn angegeben
            if (SelectedStatusFilter != null)
            {
                var statusMatches = SelectedStatusFilter.MatchesTransaction(transaction);
                filterResults.Add(statusMatches);
            }
            else if (SelectedStatus.HasValue)
            {
                var statusMatches = transaction.Status == SelectedStatus.Value;
                filterResults.Add(statusMatches);
            }

            // Betrag-Filter (wenn angegeben)
            if (MinAmount.HasValue)
            {
                var aboveMinAmount = Math.Abs(transaction.Amount) >= MinAmount.Value;
                filterResults.Add(aboveMinAmount);
            }

            if (MaxAmount.HasValue)
            {
                var belowMaxAmount = Math.Abs(transaction.Amount) <= MaxAmount.Value;
                filterResults.Add(belowMaxAmount);
            }

            // Alle angewendeten Filter müssen erfüllt sein (UND-Verknüpfung)
            if (filterResults.Count == 0)
                return true;

            return filterResults.All(result => result);
        }

        // Debug-Informationen für Filter
        public string GetActiveFiltersDebugInfo()
        {
            var activeFilters = new List<string>();

            if (!string.IsNullOrWhiteSpace(SearchText))
                activeFilters.Add($"Suche: '{SearchText}'");

            if (!string.IsNullOrWhiteSpace(SelectedCustomer))
                activeFilters.Add($"Kunde: {SelectedCustomer}");

            if (SelectedMonth.HasValue)
                activeFilters.Add($"Monat: {SelectedMonth:D2}");

            if (SelectedYear.HasValue)
                activeFilters.Add($"Jahr: {SelectedYear}");

            if (SelectedStatusFilter != null)
                activeFilters.Add($"Status: {SelectedStatusFilter.DisplayName}");

            if (MinAmount.HasValue)
                activeFilters.Add($"Min: {MinAmount:C}");

            if (MaxAmount.HasValue)
                activeFilters.Add($"Max: {MaxAmount:C}");

            return activeFilters.Count > 0
                ? string.Join(" | ", activeFilters)
                : "Keine Filter aktiv";
        }

        // Prüfung einzelner Filter
        public FilterMatchResult GetDetailedFilterResult(TransactionComparison transaction)
        {
            var result = new FilterMatchResult();

            // Freitext-Suche
            if (!string.IsNullOrWhiteSpace(SearchText))
            {
                var searchLower = SearchText.ToLowerInvariant();
                result.TextSearchMatch = transaction.CustomerEmail.ToLowerInvariant().Contains(searchLower) ||
                                       transaction.DisplayName.ToLowerInvariant().Contains(searchLower) ||
                                       transaction.StatusText.ToLowerInvariant().Contains(searchLower);
                result.HasTextSearch = true;
            }

            // Kunden-Filter
            if (!string.IsNullOrWhiteSpace(SelectedCustomer))
            {
                result.CustomerMatch = transaction.CustomerEmail.Equals(SelectedCustomer, StringComparison.OrdinalIgnoreCase);
                result.HasCustomerFilter = true;
            }

            // Monats-Filter
            if (SelectedMonth.HasValue)
            {
                result.MonthMatch = transaction.TransactionDate.Month == SelectedMonth.Value;
                result.HasMonthFilter = true;
            }

            // Jahres-Filter
            if (SelectedYear.HasValue)
            {
                result.YearMatch = transaction.TransactionDate.Year == SelectedYear.Value;
                result.HasYearFilter = true;
            }

            // Status-Filter
            if (SelectedStatusFilter != null)
            {
                result.StatusMatch = SelectedStatusFilter.MatchesTransaction(transaction);
                result.HasStatusFilter = true;
            }

            // Berechne Gesamtergebnis
            result.OverallMatch = result.AllActiveFiltersMatch();

            return result;
        }

        // INotifyPropertyChanged Implementation
        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    // Detaillierte Filter-Ergebnisse für Debugging
    public class FilterMatchResult
    {
        public bool HasTextSearch { get; set; }
        public bool TextSearchMatch { get; set; }

        public bool HasCustomerFilter { get; set; }
        public bool CustomerMatch { get; set; }

        public bool HasMonthFilter { get; set; }
        public bool MonthMatch { get; set; }

        public bool HasYearFilter { get; set; }
        public bool YearMatch { get; set; }

        public bool HasStatusFilter { get; set; }
        public bool StatusMatch { get; set; }

        public bool OverallMatch { get; set; }

        public bool AllActiveFiltersMatch()
        {
            var checks = new List<bool>();

            if (HasTextSearch) checks.Add(TextSearchMatch);
            if (HasCustomerFilter) checks.Add(CustomerMatch);
            if (HasMonthFilter) checks.Add(MonthMatch);
            if (HasYearFilter) checks.Add(YearMatch);
            if (HasStatusFilter) checks.Add(StatusMatch);

            return checks.Count == 0 || checks.All(x => x);
        }

        public string GetSummary()
        {
            var results = new List<string>();

            if (HasTextSearch) results.Add($"Text: {(TextSearchMatch ? "✓" : "✗")}");
            if (HasCustomerFilter) results.Add($"Kunde: {(CustomerMatch ? "✓" : "✗")}");
            if (HasMonthFilter) results.Add($"Monat: {(MonthMatch ? "✓" : "✗")}");
            if (HasYearFilter) results.Add($"Jahr: {(YearMatch ? "✓" : "✗")}");
            if (HasStatusFilter) results.Add($"Status: {(StatusMatch ? "✓" : "✗")}");

            return results.Count > 0
                ? $"{string.Join(", ", results)} → {(OverallMatch ? "ANZEIGEN" : "VERSTECKT")}"
                : "Keine Filter → ANZEIGEN";
        }
    }

    // Erweiterte Status-Filter-Optionen
    public class StatusFilterOption
    {
        public string DisplayName { get; set; } = string.Empty;
        public string Icon { get; set; } = string.Empty;
        public ComparisonStatus? Status { get; set; }
        public Func<TransactionComparison, bool> FilterFunc { get; set; } = _ => true;

        public bool MatchesTransaction(TransactionComparison transaction)
        {
            return FilterFunc(transaction);
        }

        // Vordefinierte Filter-Optionen
        public static StatusFilterOption All => new()
        {
            DisplayName = "Alle",
            Icon = "📋",
            Status = null,
            FilterFunc = _ => true
        };

        public static StatusFilterOption OnlyMatches => new()
        {
            DisplayName = "Nur Übereinstimmungen",
            Icon = "✅",
            Status = ComparisonStatus.Match,
            FilterFunc = t => t.Status == ComparisonStatus.Match
        };

        public static StatusFilterOption OnlyProblems => new()
        {
            DisplayName = "Nur Probleme",
            Icon = "⚠",
            Status = null,
            FilterFunc = t => t.Status != ComparisonStatus.Match
        };

        public static StatusFilterOption OnlyStripe => new()
        {
            DisplayName = "Nur in Stripe",
            Icon = "💳",
            Status = ComparisonStatus.OnlyStripe,
            FilterFunc = t => t.Status == ComparisonStatus.OnlyStripe
        };

        public static StatusFilterOption OnlyChargecloud => new()
        {
            DisplayName = "Nur in Chargecloud",
            Icon = "⚡",
            Status = ComparisonStatus.OnlyChargecloud,
            FilterFunc = t => t.Status == ComparisonStatus.OnlyChargecloud
        };

        public static StatusFilterOption AmountMismatch => new()
        {
            DisplayName = "Betragsabweichung",
            Icon = "💰",
            Status = ComparisonStatus.AmountMismatch,
            FilterFunc = t => t.Status == ComparisonStatus.AmountMismatch
        };

        // NEU: Filter für manuell bestätigte Transaktionen
        public static StatusFilterOption ManuallyConfirmed => new()
        {
            DisplayName = "Manuell bestätigt",
            Icon = "👍",
            Status = null,
            FilterFunc = t => t.MatchConfidence == MatchConfidence.Manual || t.IsManuallyConfirmed
        };

        public static List<StatusFilterOption> GetAllOptions()
        {
            return new List<StatusFilterOption>
            {
                All,
                OnlyMatches,
                OnlyProblems,
                OnlyStripe,
                OnlyChargecloud,
                AmountMismatch,
                ManuallyConfirmed // NEU
            };
        }

        public override string ToString() => DisplayName;
    }
}