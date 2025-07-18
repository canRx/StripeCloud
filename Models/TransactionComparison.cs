using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace StripeCloud.Models
{
    public class TransactionComparison : INotifyPropertyChanged
    {
        public string CustomerEmail { get; set; } = string.Empty;
        public DateTime TransactionDate { get; set; }
        public decimal Amount { get; set; }

        // Referenzen zu den Original-Transaktionen
        public StripeTransaction? StripeTransaction { get; set; }
        public ChargecloudTransaction? ChargecloudTransaction { get; set; }

        // NEUE Eigenschaften für das mehrstufige Matching
        public MatchConfidence? MatchConfidence { get; set; }
        public bool IsManuallyConfirmed { get; set; } = false;
        public bool IsManuallyRejected { get; set; } = false;

        // NEU: Property für die Checkbox-Auswahl im Edit Mode
        private bool _isSelected = false;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        // Status-Eigenschaften
        public bool HasStripeTransaction => StripeTransaction != null;
        public bool HasChargecloudTransaction => ChargecloudTransaction != null;
        public bool HasBothTransactions => HasStripeTransaction && HasChargecloudTransaction;
        public bool HasOnlyStripe => HasStripeTransaction && !HasChargecloudTransaction;
        public bool HasOnlyChargecloud => !HasStripeTransaction && HasChargecloudTransaction;

        // Diskrepanz-Erkennung
        public bool HasAmountDiscrepancy
        {
            get
            {
                if (!HasBothTransactions) return false;

                var stripeAmount = Math.Abs(StripeTransaction?.NetAmount ?? 0);
                var chargecloudAmount = Math.Abs(ChargecloudTransaction?.NetAmount ?? 0);

                // Toleranz von 0.01€ für Rundungsfehler
                return Math.Abs(stripeAmount - chargecloudAmount) > 0.01m;
            }
        }

        public decimal AmountDifference
        {
            get
            {
                if (!HasBothTransactions) return 0m;
                return (StripeTransaction?.NetAmount ?? 0) - (ChargecloudTransaction?.NetAmount ?? 0);
            }
        }

        // Status berücksichtigt jetzt auch Match-Confidence und manuelle Bestätigung
        public ComparisonStatus Status
        {
            get
            {
                if (HasOnlyStripe) return ComparisonStatus.OnlyStripe;
                if (HasOnlyChargecloud) return ComparisonStatus.OnlyChargecloud;
                if (HasAmountDiscrepancy) return ComparisonStatus.AmountMismatch;

                // Wenn es einen Match gibt, prüfe ob manuell abgelehnt
                if (HasBothTransactions && IsManuallyRejected)
                    return ComparisonStatus.ManuallyRejected;

                return ComparisonStatus.Match;
            }
        }

        // Properties für UI-Anzeige mit Match-Confidence
        public string StatusText
        {
            get
            {
                return Status switch
                {
                    ComparisonStatus.Match when MatchConfidence.HasValue => MatchConfidence.Value.GetStatusText(),
                    ComparisonStatus.Match => "✓ Übereinstimmung",
                    ComparisonStatus.OnlyStripe => "⚠ Nur in Stripe",
                    ComparisonStatus.OnlyChargecloud => "⚠ Nur in Chargecloud",
                    ComparisonStatus.AmountMismatch => "✗ Betragsabweichung",
                    ComparisonStatus.ManuallyRejected => "❌ Manuell abgelehnt",
                    _ => "Unbekannt"
                };
            }
        }

        public string StatusColor
        {
            get
            {
                return Status switch
                {
                    ComparisonStatus.Match when MatchConfidence.HasValue => MatchConfidence.Value.GetStatusColor(),
                    ComparisonStatus.Match => "#4CAF50", // Grün
                    ComparisonStatus.OnlyStripe => "#FF9800", // Orange
                    ComparisonStatus.OnlyChargecloud => "#FF9800", // Orange
                    ComparisonStatus.AmountMismatch => "#F44336", // Rot
                    ComparisonStatus.ManuallyRejected => "#9E9E9E", // Grau
                    _ => "#757575" // Grau
                };
            }
        }

        // Properties für Match-Confidence Anzeige
        public bool RequiresConfirmation => MatchConfidence?.RequiresConfirmation() == true && !IsManuallyConfirmed;
        public string MatchConfidenceText => MatchConfidence?.GetDisplayName() ?? "N/A";
        public bool ShowConfirmationButtons => RequiresConfirmation && HasBothTransactions;

        public string StripeStatus => HasStripeTransaction ? "✓" : "✗";
        public string ChargecloudStatus => HasChargecloudTransaction ? "✓" : "✗";

        public string FormattedAmount => $"{Amount:C}";
        public string FormattedDate => TransactionDate.ToString("dd.MM.yyyy");

        public string DisplayName
        {
            get
            {
                // Null-safe Zugriff
                if (HasStripeTransaction && !string.IsNullOrEmpty(StripeTransaction?.CustomerDescription))
                    return StripeTransaction.CustomerDescription;

                if (HasChargecloudTransaction && !string.IsNullOrEmpty(ChargecloudTransaction?.Contract))
                    return ChargecloudTransaction.Contract;

                return CustomerEmail;
            }
        }

        // Detailinformationen für die Detailansicht
        public string StripeDetails
        {
            get
            {
                if (!HasStripeTransaction || StripeTransaction == null)
                    return "Keine Stripe-Transaktion gefunden";

                var stripe = StripeTransaction;
                return $"ID: {stripe.Id}\n" +
                       $"Betrag: {stripe.FormattedAmount}\n" +
                       $"Status: {stripe.Status}\n" +
                       $"Erstellt: {stripe.FormattedDate}\n" +
                       $"Beschreibung: {stripe.Description}";
            }
        }

        public string ChargecloudDetails
        {
            get
            {
                if (!HasChargecloudTransaction || ChargecloudTransaction == null)
                    return "Keine Chargecloud-Transaktion gefunden";

                var cc = ChargecloudTransaction;
                return $"Rechnung: {cc.InvoiceNumber}\n" +
                       $"Betrag: {cc.FormattedAmount}\n" +
                       $"Status: {cc.PaymentStatus}\n" +
                       $"Datum: {cc.FormattedDate}\n" +
                       $"Vertrag: {cc.Contract}\n" +
                       $"Zahlungsmethode: {cc.PaymentMethod}";
            }
        }

        // Methoden für manuelle Bestätigung/Ablehnung
        public void ConfirmMatch()
        {
            IsManuallyConfirmed = true;
            IsManuallyRejected = false;
            OnPropertyChanged(nameof(RequiresConfirmation));
            OnPropertyChanged(nameof(ShowConfirmationButtons));
            OnPropertyChanged(nameof(StatusText));
            OnPropertyChanged(nameof(StatusColor));
        }

        public void RejectMatch()
        {
            IsManuallyConfirmed = false;
            IsManuallyRejected = true;
            OnPropertyChanged(nameof(Status));
            OnPropertyChanged(nameof(StatusText));
            OnPropertyChanged(nameof(StatusColor));
            OnPropertyChanged(nameof(ShowConfirmationButtons));
        }

        // INotifyPropertyChanged Implementation
        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // Factory-Methoden mit Match-Confidence
        public static TransactionComparison CreateStripeOnly(StripeTransaction stripe)
        {
            return new TransactionComparison
            {
                CustomerEmail = stripe.CustomerEmail,
                TransactionDate = stripe.CreatedDate,
                Amount = stripe.NetAmount,
                StripeTransaction = stripe
            };
        }

        public static TransactionComparison CreateChargecloudOnly(ChargecloudTransaction chargecloud)
        {
            return new TransactionComparison
            {
                CustomerEmail = chargecloud.CustomerEmail,
                TransactionDate = chargecloud.DocumentDate,
                Amount = chargecloud.NetAmount,
                ChargecloudTransaction = chargecloud
            };
        }

        public static TransactionComparison CreateMatched(StripeTransaction stripe, ChargecloudTransaction chargecloud, MatchConfidence confidence)
        {
            return new TransactionComparison
            {
                CustomerEmail = stripe.CustomerEmail,
                TransactionDate = stripe.CreatedDate,
                Amount = stripe.NetAmount,
                StripeTransaction = stripe,
                ChargecloudTransaction = chargecloud,
                MatchConfidence = confidence
            };
        }

        public override string ToString()
        {
            return $"{CustomerEmail} - {FormattedAmount} - {StatusText}";
        }
    }

    // ComparisonStatus Enum
    public enum ComparisonStatus
    {
        Match,
        OnlyStripe,
        OnlyChargecloud,
        AmountMismatch,
        ManuallyRejected
    }
}