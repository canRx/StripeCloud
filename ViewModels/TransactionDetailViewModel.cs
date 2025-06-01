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

        public ICommand CopyStripeIdCommand { get; private set; } = null!;
        public ICommand CopyChargecloudInvoiceCommand { get; private set; } = null!;
        public ICommand CopyEmailCommand { get; private set; } = null!;

        private void InitializeCommands()
        {
            CopyStripeIdCommand = new RelayCommand(() => CopyToClipboard(StripeId), () => HasStripeTransaction);
            CopyChargecloudInvoiceCommand = new RelayCommand(() => CopyToClipboard(ChargecloudInvoiceNumber), () => HasChargecloudTransaction);
            CopyEmailCommand = new RelayCommand(() => CopyToClipboard(CustomerEmail));
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
                ComparisonStatus.Match => "Beide Transaktionen stimmen überein.",
                _ => "Unbekannte Diskrepanz."
            };
        }

        private void CopyToClipboard(string text)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(text) && text != "N/A")
                {
                    System.Windows.Clipboard.SetText(text);
                }
            }
            catch
            {
                // Clipboard-Fehler ignorieren
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