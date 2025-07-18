using System;
using System.Text;
using System.Windows;
using System.Windows.Input;
using StripeCloud.Helpers;
using StripeCloud.Models;

namespace StripeCloud.ViewModels
{
    public class TransactionDetailViewModel : BaseViewModel
    {
        private TransactionComparison _comparison;

        public TransactionDetailViewModel(TransactionComparison comparison)
        {
            _comparison = comparison;
            InitializeCommands();
        }

        #region Properties

        public TransactionComparison Comparison
        {
            get => _comparison;
            set => SetProperty(ref _comparison, value);
        }

        // Allgemeine Informationen
        public string CustomerEmail => Comparison.CustomerEmail;
        public string DisplayName => Comparison.DisplayName;
        public string TransactionDate => Comparison.FormattedDate;
        public string Amount => Comparison.FormattedAmount;
        public string Status => Comparison.StatusText;
        public string StatusColor => Comparison.StatusColor;

        // NEUE Properties für Match-Confidence
        public bool HasMatchConfidence => Comparison.MatchConfidence.HasValue;
        public string MatchConfidenceText => Comparison.MatchConfidenceText;
        public bool ShowConfirmationButtons => Comparison.ShowConfirmationButtons;
        public bool RequiresConfirmation => Comparison.RequiresConfirmation;
        public string ConfidenceExplanation => GetConfidenceExplanation();

        // Stripe-spezifische Properties
        public bool HasStripeTransaction => Comparison.HasStripeTransaction;
        public StripeTransaction? StripeTransaction => Comparison.StripeTransaction;

        public string StripeId => StripeTransaction?.Id ?? "N/A";
        public string StripeAmount => StripeTransaction?.FormattedAmount ?? "N/A";
        public string StripeDate => StripeTransaction?.FormattedDate ?? "N/A";
        public string StripeStatus => StripeTransaction?.Status ?? "N/A";
        public string StripeDescription => StripeTransaction?.Description ?? "N/A";
        public string StripeCustomerId => StripeTransaction?.CustomerId ?? "N/A";
        public string StripeCustomerDescription => StripeTransaction?.CustomerDescription ?? "N/A";
        public string StripeFee => StripeTransaction?.Fee.ToString("C") ?? "N/A";
        public string StripeNetAmount => StripeTransaction?.NetAmount.ToString("C") ?? "N/A";
        public string StripeRefunded => StripeTransaction?.AmountRefunded.ToString("C") ?? "N/A";
        public bool StripeIsRefunded => StripeTransaction?.IsRefunded ?? false;
        public string StripeInvoiceId => StripeTransaction?.InvoiceId ?? "N/A";
        public string StripeCardId => StripeTransaction?.CardId ?? "N/A";
        public string StripeCurrency => StripeTransaction?.Currency?.ToUpper() ?? "N/A";

        // Chargecloud-spezifische Properties
        public bool HasChargecloudTransaction => Comparison.HasChargecloudTransaction;
        public ChargecloudTransaction? ChargecloudTransaction => Comparison.ChargecloudTransaction;

        public string ChargecloudInvoiceNumber => ChargecloudTransaction?.InvoiceNumber ?? "N/A";
        public string ChargecloudAmount => ChargecloudTransaction?.FormattedAmount ?? "N/A";
        public string ChargecloudDate => ChargecloudTransaction?.FormattedDate ?? "N/A";
        public string ChargecloudStatus => ChargecloudTransaction?.PaymentStatus ?? "N/A";
        public string ChargecloudContract => ChargecloudTransaction?.Contract ?? "N/A";
        public string ChargecloudPaymentMethod => ChargecloudTransaction?.PaymentMethod ?? "N/A";
        public string ChargecloudType => ChargecloudTransaction?.Type ?? "N/A";
        public string ChargecloudBookingNumber => ChargecloudTransaction?.BookingNumber ?? "N/A";
        public string ChargecloudServicePeriod => GetServicePeriod();
        public string ChargecloudInvoiceAmountGross => ChargecloudTransaction?.InvoiceAmountGross.ToString("C") ?? "N/A";
        public string ChargecloudPaymentAmount => ChargecloudTransaction?.PaymentAmount.ToString("C") ?? "N/A";
        public string ChargecloudOpenAmount => ChargecloudTransaction?.OpenAmount.ToString("C") ?? "N/A";
        public string ChargecloudAmountPaid => ChargecloudTransaction?.AmountPaid.ToString("C") ?? "N/A";
        public bool ChargecloudIsPaid => ChargecloudTransaction?.IsPaid ?? false;
        public bool ChargecloudIsCancelled => ChargecloudTransaction?.IsCancelled ?? false;

        // Vergleichs-spezifische Properties
        public bool HasDiscrepancy => Comparison.Status != ComparisonStatus.Match;
        public string DiscrepancyDescription => GetDiscrepancyDescription();
        public string AmountDifference => Comparison.AmountDifference.ToString("C");
        public bool HasAmountDiscrepancy => Comparison.HasAmountDiscrepancy;

        // Detailierte Beschreibungen
        public string StripeDetails => Comparison.StripeDetails;
        public string ChargecloudDetails => Comparison.ChargecloudDetails;

        #endregion

        #region Commands

        // Basis-Commands
        public ICommand CopyEmailCommand { get; private set; } = null!;

        // Stripe Copy Commands
        public ICommand CopyStripeIdCommand { get; private set; } = null!;
        public ICommand CopyStripeDescriptionCommand { get; private set; } = null!;

        // Chargecloud Copy Commands
        public ICommand CopyChargecloudInvoiceCommand { get; private set; } = null!;
        public ICommand CopyChargecloudContractCommand { get; private set; } = null!;

        // Erweiterte Commands
        public ICommand CopyAllDataCommand { get; private set; } = null!;
        public ICommand CreateEmailCommand { get; private set; } = null!;

        // NEUE Commands für Match-Confirmation
        public ICommand ConfirmMatchCommand { get; private set; } = null!;
        public ICommand RejectMatchCommand { get; private set; } = null!;

        private void InitializeCommands()
        {
            // Basis-Commands
            CopyEmailCommand = new RelayCommand(() => CopyToClipboard(CustomerEmail));

            // Stripe Copy Commands
            CopyStripeIdCommand = new RelayCommand(() => CopyToClipboard(StripeId), () => HasStripeTransaction);
            CopyStripeDescriptionCommand = new RelayCommand(() => CopyToClipboard(StripeDescription), () => HasStripeTransaction);

            // Chargecloud Copy Commands
            CopyChargecloudInvoiceCommand = new RelayCommand(() => CopyToClipboard(ChargecloudInvoiceNumber), () => HasChargecloudTransaction);
            CopyChargecloudContractCommand = new RelayCommand(() => CopyToClipboard(ChargecloudContract), () => HasChargecloudTransaction);

            // Erweiterte Commands
            CopyAllDataCommand = new RelayCommand(CopyAllData);
            CreateEmailCommand = new RelayCommand(CreateEmail);

            // NEUE Match-Confirmation Commands
            ConfirmMatchCommand = new RelayCommand(ConfirmMatch, () => ShowConfirmationButtons);
            RejectMatchCommand = new RelayCommand(RejectMatch, () => ShowConfirmationButtons);
        }

        #endregion

        #region Helper Methods

        private string GetServicePeriod()
        {
            if (ChargecloudTransaction == null)
                return "N/A";

            var start = ChargecloudTransaction.ServicePeriodStart.ToString("dd.MM.yyyy");
            var end = ChargecloudTransaction.ServicePeriodEnd.ToString("dd.MM.yyyy");

            return start == end ? start : $"{start} - {end}";
        }

        private string GetDiscrepancyDescription()
        {
            return Comparison.Status switch
            {
                ComparisonStatus.OnlyStripe => "Diese Transaktion existiert nur in Stripe, aber nicht in Chargecloud.",
                ComparisonStatus.OnlyChargecloud => "Diese Transaktion existiert nur in Chargecloud, aber nicht in Stripe.",
                ComparisonStatus.AmountMismatch => $"Die Beträge zwischen Stripe und Chargecloud weichen um {AmountDifference} ab.",
                ComparisonStatus.ManuallyRejected => "Diese Transaktionen wurden manuell als nicht zugehörig markiert.",
                ComparisonStatus.Match => "Beide Transaktionen stimmen überein.",
                _ => "Unbekannte Diskrepanz."
            };
        }

        // NEUE Methode für Confidence-Erklärung
        private string GetConfidenceExplanation()
        {
            if (!Comparison.MatchConfidence.HasValue) return "";

            return Comparison.MatchConfidence.Value switch
            {
                MatchConfidence.High => "Diese Transaktionen wurden anhand von E-Mail-Adresse und Betrag gematcht. Die Übereinstimmung ist sehr wahrscheinlich korrekt (99%).",
                MatchConfidence.Medium => "Diese Transaktionen wurden anhand des Namens (aus der E-Mail abgeleitet) und dem Betrag gematcht. Bitte prüfen Sie, ob die Zuordnung korrekt ist.",
                MatchConfidence.Low => "Diese Transaktionen wurden nur anhand des Betrags gematcht. Bitte prüfen Sie sorgfältig, ob die Zuordnung korrekt ist.",
                MatchConfidence.Manual => "Diese Transaktionen wurden manuell als zusammengehörig bestätigt.",
                _ => ""
            };
        }

        // NEUE Methoden für Match-Confirmation
        private void ConfirmMatch()
        {
            try
            {
                _comparison.ConfirmMatch();

                // UI Properties aktualisieren
                OnPropertyChanged(nameof(ShowConfirmationButtons));
                OnPropertyChanged(nameof(RequiresConfirmation));
                OnPropertyChanged(nameof(Status));
                OnPropertyChanged(nameof(StatusColor));
                OnPropertyChanged(nameof(ConfidenceExplanation));

                MessageBox.Show("Die Transaktionszuordnung wurde bestätigt.", "Bestätigt",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Bestätigen: {ex.Message}", "Fehler",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RejectMatch()
        {
            try
            {
                var result = MessageBox.Show(
                    "Möchten Sie die Transaktionszuordnung wirklich aufheben?\n\n" +
                    "Die Transaktionen werden dann als separate Einträge angezeigt.",
                    "Zuordnung aufheben",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    _comparison.RejectMatch();

                    // UI Properties aktualisieren
                    OnPropertyChanged(nameof(ShowConfirmationButtons));
                    OnPropertyChanged(nameof(Status));
                    OnPropertyChanged(nameof(StatusColor));
                    OnPropertyChanged(nameof(HasDiscrepancy));
                    OnPropertyChanged(nameof(DiscrepancyDescription));

                    MessageBox.Show("Die Transaktionszuordnung wurde aufgehoben.", "Aufgehoben",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Aufheben: {ex.Message}", "Fehler",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CopyToClipboard(string text)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(text) && text != "N/A")
                {
                    Clipboard.SetText(text);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Kopieren: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CopyAllData()
        {
            try
            {
                var sb = new StringBuilder();

                // Header
                sb.AppendLine("=== TRANSAKTIONSDETAILS ===");
                sb.AppendLine($"Kunde: {CustomerEmail}");
                sb.AppendLine($"Datum: {TransactionDate}");
                sb.AppendLine($"Status: {Status}");

                // Match-Confidence Info
                if (HasMatchConfidence)
                {
                    sb.AppendLine($"Match-Level: {MatchConfidenceText}");
                    sb.AppendLine($"Erklärung: {ConfidenceExplanation}");
                }

                sb.AppendLine();

                // Stripe-Daten
                sb.AppendLine("=== STRIPE ===");
                if (HasStripeTransaction)
                {
                    sb.AppendLine($"ID: {StripeId}");
                    sb.AppendLine($"Betrag: {StripeAmount}");
                    sb.AppendLine($"Status: {StripeStatus}");
                    sb.AppendLine($"Erstellt: {StripeDate}");
                    sb.AppendLine($"Beschreibung: {StripeDescription}");
                    sb.AppendLine($"Gebühr: {StripeFee}");
                    sb.AppendLine($"Nettobetrag: {StripeNetAmount}");
                    sb.AppendLine($"Rückerstattet: {StripeRefunded}");
                }
                else
                {
                    sb.AppendLine("Keine Stripe-Transaktion gefunden");
                }
                sb.AppendLine();

                // Chargecloud-Daten
                sb.AppendLine("=== CHARGECLOUD ===");
                if (HasChargecloudTransaction)
                {
                    sb.AppendLine($"Rechnungsnummer: {ChargecloudInvoiceNumber}");
                    sb.AppendLine($"Betrag: {ChargecloudAmount}");
                    sb.AppendLine($"Status: {ChargecloudStatus}");
                    sb.AppendLine($"Datum: {ChargecloudDate}");
                    sb.AppendLine($"Vertrag: {ChargecloudContract}");
                    sb.AppendLine($"Zahlungsmethode: {ChargecloudPaymentMethod}");
                    sb.AppendLine($"Bruttobetrag: {ChargecloudInvoiceAmountGross}");
                    sb.AppendLine($"Bezahlt: {ChargecloudAmountPaid}");
                    sb.AppendLine($"Offen: {ChargecloudOpenAmount}");
                }
                else
                {
                    sb.AppendLine("Keine Chargecloud-Transaktion gefunden");
                }
                sb.AppendLine();

                // Vergleichsanalyse
                if (HasDiscrepancy)
                {
                    sb.AppendLine("=== VERGLEICHSANALYSE ===");
                    sb.AppendLine($"Problem: {DiscrepancyDescription}");
                    sb.AppendLine($"Betragsunterschied: {AmountDifference}");
                    sb.AppendLine($"Match-Score: {MatchingScore}");
                    sb.AppendLine($"Timeline-Analyse: {TimelineAnalysis}");
                    sb.AppendLine($"Betrag-Analyse: {AmountAnalysis}");
                }

                Clipboard.SetText(sb.ToString());
                MessageBox.Show("Alle Transaktionsdaten wurden in die Zwischenablage kopiert.", "Kopiert", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Kopieren aller Daten: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CreateEmail()
        {
            try
            {
                var subject = $"Transaktionsproblem: {CustomerEmail} - {Amount}";
                var body = new StringBuilder();

                body.AppendLine($"Sehr geehrte Damen und Herren,");
                body.AppendLine();
                body.AppendLine($"bezüglich der Transaktion für {CustomerEmail} vom {TransactionDate} liegt folgendes Problem vor:");
                body.AppendLine();
                body.AppendLine($"Problem: {DiscrepancyDescription}");

                if (HasMatchConfidence)
                {
                    body.AppendLine($"Match-Level: {MatchConfidenceText}");
                }

                body.AppendLine();

                if (HasStripeTransaction)
                {
                    body.AppendLine("Stripe-Details:");
                    body.AppendLine($"- ID: {StripeId}");
                    body.AppendLine($"- Betrag: {StripeAmount}");
                    body.AppendLine($"- Status: {StripeStatus}");
                    body.AppendLine();
                }

                if (HasChargecloudTransaction)
                {
                    body.AppendLine("Chargecloud-Details:");
                    body.AppendLine($"- Rechnungsnummer: {ChargecloudInvoiceNumber}");
                    body.AppendLine($"- Betrag: {ChargecloudAmount}");
                    body.AppendLine($"- Status: {ChargecloudStatus}");
                    body.AppendLine();
                }

                body.AppendLine("Bitte prüfen Sie diesen Fall und geben Sie eine Rückmeldung.");
                body.AppendLine();
                body.AppendLine("Mit freundlichen Grüßen");

                var mailtoUri = $"mailto:{CustomerEmail}?subject={Uri.EscapeDataString(subject)}&body={Uri.EscapeDataString(body.ToString())}";

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = mailtoUri,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Erstellen der E-Mail: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Additional Analysis

        // Zusätzliche Analyse-Properties für Power-User
        public string MatchingScore => CalculateMatchingScore();
        public string TimelineAnalysis => GenerateTimelineAnalysis();
        public string AmountAnalysis => GenerateAmountAnalysis();

        private string CalculateMatchingScore()
        {
            if (!Comparison.HasBothTransactions)
                return "N/A - Keine vollständige Übereinstimmung möglich";

            var stripe = Comparison.StripeTransaction!;
            var chargecloud = Comparison.ChargecloudTransaction!;

            // Datum-Übereinstimmung
            var dateDiff = Math.Abs((stripe.CreatedDate - chargecloud.DocumentDate).TotalDays);
            var dateScore = Math.Max(0, 100 - (dateDiff * 10)); // 10% Abzug pro Tag Differenz

            // Betrag-Übereinstimmung
            var amountDiff = Math.Abs(stripe.NetAmount - chargecloud.NetAmount);
            var amountScore = amountDiff <= 0.01m ? 100 : Math.Max(0, 100 - (double)(amountDiff * 10));

            var overallScore = (dateScore + amountScore) / 2;

            return $"{overallScore:F1}% (Datum: {dateScore:F1}%, Betrag: {amountScore:F1}%)";
        }

        private string GenerateTimelineAnalysis()
        {
            if (!Comparison.HasBothTransactions)
                return "Keine Timeline-Analyse möglich - nur eine Transaktion vorhanden";

            var stripe = Comparison.StripeTransaction!;
            var chargecloud = Comparison.ChargecloudTransaction!;

            var timeDiff = (stripe.CreatedDate - chargecloud.DocumentDate).TotalDays;

            if (Math.Abs(timeDiff) <= 1)
                return "✓ Transaktionen fanden am gleichen Tag statt";
            else if (timeDiff > 0)
                return $"⚠ Stripe-Transaktion {Math.Abs(timeDiff):F0} Tage nach Chargecloud-Transaktion";
            else
                return $"⚠ Chargecloud-Transaktion {Math.Abs(timeDiff):F0} Tage nach Stripe-Transaktion";
        }

        private string GenerateAmountAnalysis()
        {
            if (!Comparison.HasBothTransactions)
                return Comparison.HasStripeTransaction
                    ? $"Nur Stripe: {StripeAmount}"
                    : $"Nur Chargecloud: {ChargecloudAmount}";

            var stripe = Comparison.StripeTransaction!;
            var chargecloud = Comparison.ChargecloudTransaction!;

            if (Comparison.HasAmountDiscrepancy)
            {
                var diff = stripe.NetAmount - chargecloud.NetAmount;
                var percentage = chargecloud.NetAmount != 0 ? (diff / chargecloud.NetAmount * 100) : 0;

                return $"Differenz: {diff:C} ({percentage:+0.0;-0.0}%) - " +
                       $"Stripe: {stripe.NetAmount:C}, Chargecloud: {chargecloud.NetAmount:C}";
            }
            else
            {
                return $"✓ Beträge stimmen überein: {stripe.NetAmount:C}";
            }
        }

        #endregion
    }
}