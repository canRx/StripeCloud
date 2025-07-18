using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Win32;
using StripeCloud.Helpers;
using StripeCloud.Models;
using StripeCloud.Services;

namespace StripeCloud.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private readonly CsvImportService _csvImportService;
        private readonly TransactionComparisonService _comparisonService;
        private readonly ExcelExportService _excelExportService;
        private readonly ManualMatchingService _manualMatchingService;

        // Private fields
        private bool _isLoading;
        private string _statusMessage = "Bereit";
        private TransactionComparison? _selectedTransaction;
        private FilterOptions _filterOptions;
        private ComparisonResult? _lastComparisonResult;

        // Bearbeitungsmodus
        private bool _isEditMode;
        private ObservableCollection<TransactionComparison> _selectedTransactions = new();

        // Collections
        private List<StripeTransaction> _stripeTransactions = new();
        private List<ChargecloudTransaction> _chargecloudTransactions = new();
        private List<TransactionComparison> _allComparisons = new();
        private ObservableCollection<TransactionComparison> _filteredComparisons = new();
        private ObservableCollection<string> _availableCustomers = new();
        private ObservableCollection<int> _availableYears = new();
        private ObservableCollection<int> _availableMonths = new();
        private ObservableCollection<StatusFilterOption> _availableStatusFilters = new();

        // Sortierung
        private ICollectionView? _filteredComparisonsView;
        private string _currentSortColumn = string.Empty;
        private ListSortDirection _currentSortDirection = ListSortDirection.Ascending;

        public MainViewModel()
        {
            _csvImportService = new CsvImportService();
            _comparisonService = new TransactionComparisonService();
            _excelExportService = new ExcelExportService();
            _manualMatchingService = new ManualMatchingService();
            _filterOptions = new FilterOptions();

            InitializeCommands();
            InitializeData();
            InitializeCollectionView();

            // Filter-Changes überwachen
            _filterOptions.PropertyChanged += OnFilterChanged;
        }

        #region Properties

        public bool IsLoading
        {
            get => _isLoading;
            set => SetProperty(ref _isLoading, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public TransactionComparison? SelectedTransaction
        {
            get => _selectedTransaction;
            set => SetProperty(ref _selectedTransaction, value);
        }

        public FilterOptions FilterOptions
        {
            get => _filterOptions;
            set => SetProperty(ref _filterOptions, value);
        }

        public ObservableCollection<TransactionComparison> FilteredComparisons
        {
            get => _filteredComparisons;
            set => SetProperty(ref _filteredComparisons, value);
        }

        // Bearbeitungsmodus Properties
        public bool IsEditMode
        {
            get => _isEditMode;
            set
            {
                if (SetProperty(ref _isEditMode, value))
                {
                    if (!value)
                    {
                        // Beim Verlassen des Bearbeitungsmodus: Auswahl zurücksetzen
                        foreach (var transaction in SelectedTransactions.ToList())
                        {
                            transaction.IsSelected = false;
                        }
                        SelectedTransactions.Clear();
                    }
                    OnPropertyChanged(nameof(EditModeStatusText));
                    OnPropertyChanged(nameof(SelectedTransactionCount));
                    OnPropertyChanged(nameof(CanManualMatch));
                    OnPropertyChanged(nameof(CanUnmatch));
                }
            }
        }

        public ObservableCollection<TransactionComparison> SelectedTransactions
        {
            get => _selectedTransactions;
            set => SetProperty(ref _selectedTransactions, value);
        }

        public string EditModeStatusText
        {
            get
            {
                if (!IsEditMode)
                    return "Bearbeitungsmodus inaktiv";

                return $"Bearbeitungsmodus aktiv - {SelectedTransactionCount} Transaktionen ausgewählt";
            }
        }

        public int SelectedTransactionCount => SelectedTransactions.Count;

        public bool CanManualMatch
        {
            get
            {
                if (!IsEditMode || SelectedTransactionCount < 2)
                    return false;

                // Mindestens eine Stripe und eine Chargecloud Transaktion
                var hasStripe = SelectedTransactions.Any(t => t.HasStripeTransaction);
                var hasChargecloud = SelectedTransactions.Any(t => t.HasChargecloudTransaction);

                return hasStripe && hasChargecloud;
            }
        }

        public bool CanUnmatch
        {
            get
            {
                if (!IsEditMode || SelectedTransactionCount == 0)
                    return false;

                // Alle ausgewählten Transaktionen müssen gematchte Transaktionen sein
                return SelectedTransactions.All(t => t.Status == ComparisonStatus.Match && t.HasBothTransactions);
            }
        }

        // CollectionView für Sortierung
        public ICollectionView? FilteredComparisonsView
        {
            get => _filteredComparisonsView;
            set => SetProperty(ref _filteredComparisonsView, value);
        }

        public ObservableCollection<string> AvailableCustomers
        {
            get => _availableCustomers;
            set => SetProperty(ref _availableCustomers, value);
        }

        public ObservableCollection<int> AvailableYears
        {
            get => _availableYears;
            set => SetProperty(ref _availableYears, value);
        }

        public ObservableCollection<int> AvailableMonths
        {
            get => _availableMonths;
            set => SetProperty(ref _availableMonths, value);
        }

        public ObservableCollection<StatusFilterOption> AvailableStatusFilters
        {
            get => _availableStatusFilters;
            set => SetProperty(ref _availableStatusFilters, value);
        }

        // Statistik-Properties
        public int TotalTransactions => _allComparisons.Count;
        public int PerfectMatches => _allComparisons.Count(c => c.Status == ComparisonStatus.Match);
        public int ProblemsCount => _allComparisons.Count(c => c.Status != ComparisonStatus.Match);

        public string DisplayStatusInfo =>
            $"Angezeigt: {FilteredComparisons.Count} von {_allComparisons.Count}";

        public bool HasData => _allComparisons.Count > 0;
        public bool HasStripeData => _stripeTransactions.Count > 0;
        public bool HasChargecloudData => _chargecloudTransactions.Count > 0;

        // Status-Karten Properties
        public int StripeTransactionCount => _stripeTransactions.Count;
        public int ChargecloudTransactionCount => _chargecloudTransactions.Count;

        // Sortier-Info
        public string CurrentSortInfo
        {
            get
            {
                if (string.IsNullOrEmpty(_currentSortColumn))
                    return "Sortierung: Standard";

                var direction = _currentSortDirection == ListSortDirection.Ascending ? "↑" : "↓";
                return $"Sortierung: {GetColumnDisplayName(_currentSortColumn)} {direction}";
            }
        }

        #endregion

        #region Commands

        public ICommand LoadStripeFileCommand { get; private set; } = null!;
        public ICommand LoadChargecloudFileCommand { get; private set; } = null!;
        public ICommand RefreshDataCommand { get; private set; } = null!;
        public ICommand ClearFiltersCommand { get; private set; } = null!;
        public ICommand ExportToExcelCommand { get; private set; } = null!;
        public ICommand ShowTransactionDetailsCommand { get; private set; } = null!;
        public ICommand OpenDetailWindowCommand { get; private set; } = null!;
        public ICommand OpenSupportEmailCommand { get; private set; } = null!;

        // Filter-Commands
        public ICommand ShowOnlyMatchesCommand { get; private set; } = null!;
        public ICommand ShowOnlyProblemsCommand { get; private set; } = null!;
        public ICommand ShowOnlyAmountMismatchCommand { get; private set; } = null!;
        public ICommand ShowAllTransactionsCommand { get; private set; } = null!;
        public ICommand ShowManuallyConfirmedCommand { get; private set; } = null!;

        // Bearbeitungsmodus Commands
        public ICommand ToggleEditModeCommand { get; private set; } = null!;
        public ICommand ToggleTransactionSelectionCommand { get; private set; } = null!;
        public ICommand ManualMatchCommand { get; private set; } = null!;
        public ICommand UnmatchCommand { get; private set; } = null!;
        public ICommand ClearSelectionCommand { get; private set; } = null!;

        private void InitializeCommands()
        {
            LoadStripeFileCommand = new RelayCommand(async () => await LoadStripeFileAsync(), () => !IsLoading);
            LoadChargecloudFileCommand = new RelayCommand(async () => await LoadChargecloudFileAsync(), () => !IsLoading);
            RefreshDataCommand = new RelayCommand(async () => await RefreshComparisonAsync(), () => !IsLoading && HasStripeData && HasChargecloudData);
            ClearFiltersCommand = new RelayCommand(ClearAllFilters, () => FilterOptions.HasAnyFilter);
            ExportToExcelCommand = new RelayCommand(async () => await ExportToExcelAsync(), () => HasData && !IsLoading);
            ShowTransactionDetailsCommand = new RelayCommand<TransactionComparison>(ShowTransactionDetails);
            OpenDetailWindowCommand = new RelayCommand(OpenDetailWindow, () => SelectedTransaction != null);
            OpenSupportEmailCommand = new RelayCommand(OpenSupportEmail);

            // Schnellfilter-Commands
            ShowOnlyMatchesCommand = new RelayCommand(ShowOnlyMatches, () => HasData);
            ShowOnlyProblemsCommand = new RelayCommand(ShowOnlyProblems, () => HasData);
            ShowOnlyAmountMismatchCommand = new RelayCommand(ShowOnlyAmountMismatch, () => HasData);
            ShowAllTransactionsCommand = new RelayCommand(ShowAllTransactions, () => HasData);
            ShowManuallyConfirmedCommand = new RelayCommand(ShowManuallyConfirmed, () => HasData);

            // Bearbeitungsmodus Commands
            ToggleEditModeCommand = new RelayCommand(ToggleEditMode, () => HasData);
            ToggleTransactionSelectionCommand = new RelayCommand<TransactionComparison>(ToggleTransactionSelection);
            ManualMatchCommand = new RelayCommand(ManualMatch, () => CanManualMatch);
            UnmatchCommand = new RelayCommand(Unmatch, () => CanUnmatch);
            ClearSelectionCommand = new RelayCommand(ClearSelection, () => SelectedTransactionCount > 0);

            // Event-Handler für SelectedTransactions-Collection
            SelectedTransactions.CollectionChanged += (s, e) =>
            {
                OnPropertyChanged(nameof(SelectedTransactionCount));
                OnPropertyChanged(nameof(EditModeStatusText));
                OnPropertyChanged(nameof(CanManualMatch));
                OnPropertyChanged(nameof(CanUnmatch));
            };
        }

        #endregion

        #region Methods

        private void InitializeData()
        {
            // Standard-Monate hinzufügen
            for (int i = 1; i <= 12; i++)
            {
                AvailableMonths.Add(i);
            }

            // Standard-Jahre hinzufügen (aktuelles Jahr ± 2)
            var currentYear = DateTime.Now.Year;
            for (int i = currentYear - 2; i <= currentYear + 1; i++)
            {
                AvailableYears.Add(i);
            }

            // Status-Filter-Optionen hinzufügen
            foreach (var option in StatusFilterOption.GetAllOptions())
            {
                AvailableStatusFilters.Add(option);
            }
        }

        private void InitializeCollectionView()
        {
            FilteredComparisonsView = CollectionViewSource.GetDefaultView(FilteredComparisons);
            FilteredComparisonsView.SortDescriptions.Clear();
        }

        // Helper-Methode für Checkbox-Binding
        public bool IsTransactionSelected(TransactionComparison transaction)
        {
            return SelectedTransactions.Contains(transaction);
        }

        // Bearbeitungsmodus Methods
        private void ToggleEditMode()
        {
            IsEditMode = !IsEditMode;

            if (IsEditMode)
            {
                StatusMessage = "Bearbeitungsmodus aktiviert - Wählen Sie Transaktionen zum Matchen aus";
            }
            else
            {
                StatusMessage = "Bearbeitungsmodus deaktiviert";
            }
        }

        private void ToggleTransactionSelection(TransactionComparison? transaction)
        {
            if (!IsEditMode || transaction == null)
                return;

            if (SelectedTransactions.Contains(transaction))
            {
                SelectedTransactions.Remove(transaction);
                transaction.IsSelected = false;
            }
            else
            {
                SelectedTransactions.Add(transaction);
                transaction.IsSelected = true;
            }
        }

        private void ManualMatch()
        {
            if (!CanManualMatch)
                return;

            try
            {
                IsLoading = true;
                StatusMessage = "Führe manuelles Matching durch...";

                // Zeige Confirmation-Dialog mit Details
                var confirmationResult = ShowMatchingConfirmationDialog();
                if (confirmationResult != MessageBoxResult.Yes)
                {
                    StatusMessage = "Manuelles Matching abgebrochen";
                    return;
                }

                // Führe das manuelle Matching durch
                var result = _manualMatchingService.MatchTransactions(_allComparisons, SelectedTransactions.ToList());

                if (result.IsSuccessful)
                {
                    // Aktualisiere die Daten
                    _allComparisons = result.UpdatedComparisons;
                    UpdateAvailableFilterOptions();
                    ApplyFilters();

                    // Bearbeitungsmodus verlassen
                    IsEditMode = false;

                    StatusMessage = $"Manuelles Matching erfolgreich: {result.MatchedCount} Transaktionen verknüpft";

                    MessageBox.Show($"Manuelles Matching erfolgreich abgeschlossen!\n\n" +
                                  $"• {result.MatchedCount} Transaktionen wurden verknüpft\n" +
                                  $"• Die neuen Matches sind als 'Manuell bestätigt' markiert",
                                  "Matching erfolgreich", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    StatusMessage = $"Fehler beim manuellen Matching: {result.ErrorMessage}";
                    MessageBox.Show($"Fehler beim manuellen Matching:\n\n{result.ErrorMessage}",
                                  "Matching-Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Fehler beim manuellen Matching: {ex.Message}";
                MessageBox.Show($"Unerwarteter Fehler beim manuellen Matching:\n\n{ex.Message}",
                              "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
                UpdateCommandStates();
                NotifyStatisticsChanged();
            }
        }

        private void Unmatch()
        {
            if (!CanUnmatch)
                return;

            try
            {
                IsLoading = true;
                StatusMessage = "Hebe Matching auf...";

                // Confirmation-Dialog
                var result = MessageBox.Show(
                    $"Möchten Sie das Matching von {SelectedTransactionCount} Transaktionen wirklich aufheben?\n\n" +
                    "Die Transaktionen werden wieder als separate Einträge angezeigt.",
                    "Matching aufheben",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result != MessageBoxResult.Yes)
                {
                    StatusMessage = "Aufheben des Matchings abgebrochen";
                    return;
                }

                // Führe das Unmatch durch
                var unmatchResult = _manualMatchingService.UnmatchTransactions(_allComparisons, SelectedTransactions.ToList());

                if (unmatchResult.IsSuccessful)
                {
                    // Aktualisiere die Daten
                    _allComparisons = unmatchResult.UpdatedComparisons;
                    UpdateAvailableFilterOptions();
                    ApplyFilters();

                    // Bearbeitungsmodus verlassen
                    IsEditMode = false;

                    StatusMessage = $"Matching aufgehoben: {unmatchResult.UnmatchedCount} Transaktionen getrennt";

                    MessageBox.Show($"Matching erfolgreich aufgehoben!\n\n" +
                                  $"• {unmatchResult.UnmatchedCount} Transaktionen wurden getrennt\n" +
                                  $"• Die Transaktionen sind nun wieder einzeln verfügbar",
                                  "Matching aufgehoben", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    StatusMessage = $"Fehler beim Aufheben: {unmatchResult.ErrorMessage}";
                    MessageBox.Show($"Fehler beim Aufheben des Matchings:\n\n{unmatchResult.ErrorMessage}",
                                  "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Fehler beim Aufheben: {ex.Message}";
                MessageBox.Show($"Unerwarteter Fehler beim Aufheben des Matchings:\n\n{ex.Message}",
                              "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
                UpdateCommandStates();
                NotifyStatisticsChanged();
            }
        }

        private void ClearSelection()
        {
            foreach (var transaction in SelectedTransactions.ToList())
            {
                transaction.IsSelected = false;
            }
            SelectedTransactions.Clear();
            StatusMessage = "Auswahl zurückgesetzt";
        }

        private MessageBoxResult ShowMatchingConfirmationDialog()
        {
            var stripeCount = SelectedTransactions.Count(t => t.HasStripeTransaction);
            var chargecloudCount = SelectedTransactions.Count(t => t.HasChargecloudTransaction);

            var message = $"Manuelles Matching durchführen?\n\n" +
                         $"Ausgewählte Transaktionen:\n" +
                         $"• {stripeCount} Stripe-Transaktionen\n" +
                         $"• {chargecloudCount} Chargecloud-Transaktionen\n\n" +
                         $"Es werden {Math.Min(stripeCount, chargecloudCount)} Matches erstellt.\n";

            if (stripeCount != chargecloudCount)
            {
                message += $"\n⚠️ Warnung: Ungleiche Anzahl - {Math.Abs(stripeCount - chargecloudCount)} Transaktionen bleiben unverknüpft.";
            }

            return MessageBox.Show(message, "Manuelles Matching bestätigen",
                                 MessageBoxButton.YesNo, MessageBoxImage.Question);
        }

        // Support E-Mail öffnen
        private void OpenSupportEmail()
        {
            try
            {
                var subject = "StripeCloud Support - Anfrage";
                var body = "Hallo Support-Team,\n\nIch habe eine Frage/ein Problem mit StripeCloud:\n\n[Bitte beschreiben Sie hier Ihr Anliegen]\n\nVielen Dank für Ihre Hilfe!\n\nMit freundlichen Grüßen";

                var mailtoUri = $"mailto:caner.engin@smartwerk.org?subject={Uri.EscapeDataString(subject)}&body={Uri.EscapeDataString(body)}";

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = mailtoUri,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"E-Mail-Client konnte nicht geöffnet werden.\nBitte wenden Sie sich direkt an: caner.engin@smartwerk.org\n\nFehler: {ex.Message}",
                    "Support", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // Sortier-Event Handler
        public void HandleDataGridSorting(string? columnName, ListSortDirection direction)
        {
            if (string.IsNullOrEmpty(columnName) || FilteredComparisonsView == null)
                return;

            _currentSortColumn = columnName;
            _currentSortDirection = direction;

            FilteredComparisonsView.SortDescriptions.Clear();
            FilteredComparisonsView.SortDescriptions.Add(new SortDescription(columnName, direction));

            OnPropertyChanged(nameof(CurrentSortInfo));
        }

        private string GetColumnDisplayName(string columnName)
        {
            return columnName switch
            {
                "CustomerEmail" => "Kunde",
                "TransactionDate" => "Datum",
                "Amount" => "Betrag",
                "Status" => "Status",
                "HasStripeTransaction" => "Stripe",
                "HasChargecloudTransaction" => "Chargecloud",
                _ => columnName
            };
        }

        private async Task LoadStripeFileAsync()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Stripe CSV-Datei auswählen",
                Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*",
                FilterIndex = 1
            };

            if (openFileDialog.ShowDialog() == true)
            {
                await LoadStripeFileAsync(openFileDialog.FileName);
            }
        }

        private async Task LoadStripeFileAsync(string filePath)
        {
            try
            {
                IsLoading = true;
                StatusMessage = "Lade Stripe-Daten...";

                var result = await _csvImportService.ImportStripeWithStatsAsync(filePath);

                if (result.IsSuccessful)
                {
                    _stripeTransactions = result.Data;
                    StatusMessage = $"Stripe-Daten geladen: {result.SuccessfulRows} Transaktionen in {result.ImportDuration.TotalSeconds:F1}s";

                    if (HasChargecloudData)
                    {
                        await RefreshComparisonAsync();
                    }
                }
                else
                {
                    StatusMessage = $"Fehler beim Laden der Stripe-Daten: {string.Join(", ", result.Errors)}";
                    MessageBox.Show($"Fehler beim Importieren:\n{string.Join("\n", result.Errors)}",
                        "Import-Fehler", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Fehler: {ex.Message}";
                MessageBox.Show($"Fehler beim Laden der Stripe-Datei:\n{ex.Message}",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
                UpdateCommandStates();
                NotifyStatisticsChanged();
            }
        }

        private async Task LoadChargecloudFileAsync()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Chargecloud CSV-Datei auswählen",
                Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*",
                FilterIndex = 1
            };

            if (openFileDialog.ShowDialog() == true)
            {
                await LoadChargecloudFileAsync(openFileDialog.FileName);
            }
        }

        private async Task LoadChargecloudFileAsync(string filePath)
        {
            try
            {
                IsLoading = true;
                StatusMessage = "Lade Chargecloud-Daten...";

                var result = await _csvImportService.ImportChargecloudWithStatsAsync(filePath);

                if (result.IsSuccessful)
                {
                    _chargecloudTransactions = result.Data;
                    StatusMessage = $"Chargecloud-Daten geladen: {result.SuccessfulRows} Transaktionen in {result.ImportDuration.TotalSeconds:F1}s";

                    if (HasStripeData)
                    {
                        await RefreshComparisonAsync();
                    }
                }
                else
                {
                    StatusMessage = $"Fehler beim Laden der Chargecloud-Daten: {string.Join(", ", result.Errors)}";
                    MessageBox.Show($"Fehler beim Importieren:\n{string.Join("\n", result.Errors)}",
                        "Import-Fehler", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Fehler: {ex.Message}";
                MessageBox.Show($"Fehler beim Laden der Chargecloud-Datei:\n{ex.Message}",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
                UpdateCommandStates();
                NotifyStatisticsChanged();
            }
        }

        private async Task RefreshComparisonAsync()
        {
            if (!HasStripeData || !HasChargecloudData)
                return;

            try
            {
                IsLoading = true;
                StatusMessage = "Vergleiche Transaktionen...";

                // Bearbeitungsmodus deaktivieren beim Refresh
                IsEditMode = false;

                _lastComparisonResult = await _comparisonService.CompareTransactionsAsync(_stripeTransactions, _chargecloudTransactions);

                if (_lastComparisonResult.IsSuccessful)
                {
                    _allComparisons = _lastComparisonResult.Comparisons;
                    UpdateAvailableFilterOptions();
                    ApplyFilters();

                    StatusMessage = $"Vergleich abgeschlossen: {TotalTransactions} Transaktionen, {PerfectMatches} Übereinstimmungen, {ProblemsCount} Probleme";
                }
                else
                {
                    StatusMessage = $"Vergleichsfehler: {string.Join(", ", _lastComparisonResult.Errors)}";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Fehler beim Vergleich: {ex.Message}";
                MessageBox.Show($"Fehler beim Transaktionsvergleich:\n{ex.Message}",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
                UpdateCommandStates();
                NotifyStatisticsChanged();
            }
        }

        private void UpdateAvailableFilterOptions()
        {
            // Verfügbare Kunden aktualisieren
            var customers = _allComparisons
                .Select(c => c.CustomerEmail)
                .Distinct()
                .OrderBy(email => email)
                .ToList();

            AvailableCustomers.Clear();
            foreach (var customer in customers)
            {
                AvailableCustomers.Add(customer);
            }

            // Verfügbare Jahre aktualisieren
            var years = _allComparisons
                .Select(c => c.TransactionDate.Year)
                .Distinct()
                .OrderBy(year => year)
                .ToList();

            AvailableYears.Clear();
            foreach (var year in years)
            {
                AvailableYears.Add(year);
            }
        }

        private void OnFilterChanged(object? sender, PropertyChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void ApplyFilters()
        {
            var filtered = _comparisonService.ApplyFilters(_allComparisons, FilterOptions);

            FilteredComparisons.Clear();
            foreach (var comparison in filtered)
            {
                FilteredComparisons.Add(comparison);
            }

            // CollectionView refresh
            FilteredComparisonsView?.Refresh();

            OnPropertyChanged(nameof(DisplayStatusInfo));
        }

        private void ClearAllFilters()
        {
            FilterOptions.ClearAllFilters();
        }

        // Schnellfilter-Methoden
        private void ShowOnlyMatches()
        {
            FilterOptions.SetOnlyMatches();
        }

        private void ShowOnlyProblems()
        {
            FilterOptions.SetOnlyProblems();
        }

        private void ShowOnlyAmountMismatch()
        {
            FilterOptions.SetOnlyAmountMismatch();
        }

        private void ShowManuallyConfirmed()
        {
            FilterOptions.SetManuallyConfirmed();
        }

        private void ShowAllTransactions()
        {
            FilterOptions.ShowAll();
        }

        private async Task ExportToExcelAsync()
        {
            if (_lastComparisonResult == null)
                return;

            try
            {
                IsLoading = true;
                StatusMessage = "Exportiere nach Excel...";

                var saveFileDialog = new SaveFileDialog
                {
                    Title = "Excel-Export speichern",
                    Filter = "Excel-Dateien (*.xlsx)|*.xlsx",
                    FileName = $"Transaktionsvergleich_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    var filePath = await _excelExportService.ExportComparisonResultsAsync(
                        FilteredComparisons.ToList(),
                        _lastComparisonResult,
                        saveFileDialog.FileName);

                    StatusMessage = $"Export erfolgreich: {filePath}";

                    var result = MessageBox.Show($"Export erfolgreich erstellt:\n{filePath}\n\nDatei öffnen?",
                        "Export erfolgreich", MessageBoxButton.YesNo, MessageBoxImage.Information);

                    if (result == MessageBoxResult.Yes)
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = filePath,
                            UseShellExecute = true
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Export-Fehler: {ex.Message}";
                MessageBox.Show($"Fehler beim Excel-Export:\n{ex.Message}",
                    "Export-Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void ShowTransactionDetails(TransactionComparison? transaction)
        {
            if (transaction != null)
            {
                SelectedTransaction = transaction;
            }
        }

        private void OpenDetailWindow()
        {
            if (SelectedTransaction != null)
            {
                var detailViewModel = new TransactionDetailViewModel(SelectedTransaction);
                var detailWindow = new Views.TransactionDetailWindow
                {
                    DataContext = detailViewModel,
                    Owner = Application.Current.MainWindow
                };
                detailWindow.Show();
            }
        }

        private void UpdateCommandStates()
        {
            CommandManager.InvalidateRequerySuggested();
        }

        private void NotifyStatisticsChanged()
        {
            OnPropertyChanged(nameof(HasData));
            OnPropertyChanged(nameof(HasStripeData));
            OnPropertyChanged(nameof(HasChargecloudData));
            OnPropertyChanged(nameof(StripeTransactionCount));
            OnPropertyChanged(nameof(ChargecloudTransactionCount));
            OnPropertyChanged(nameof(TotalTransactions));
            OnPropertyChanged(nameof(PerfectMatches));
            OnPropertyChanged(nameof(ProblemsCount));
            OnPropertyChanged(nameof(DisplayStatusInfo));
        }

        #endregion
    }
}