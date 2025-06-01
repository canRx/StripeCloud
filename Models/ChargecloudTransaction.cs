using System;
using CsvHelper.Configuration.Attributes;
using StripeCloud.Helpers;

namespace StripeCloud.Models
{
    public class ChargecloudTransaction
    {
        [Name("#")]
        public int Number { get; set; }

        [Name("Rechnungsnummer")]
        public string InvoiceNumber { get; set; } = string.Empty;

        [Name("Typ")]
        public string Type { get; set; } = string.Empty;

        [Name("Leistungszeitraum (Start)")]
        [TypeConverter(typeof(GermanDateTimeConverter))]
        public DateTime ServicePeriodStart { get; set; }

        [Name("Leistungszeitraum (Ende)")]
        [TypeConverter(typeof(GermanDateTimeConverter))]
        public DateTime ServicePeriodEnd { get; set; }

        [Name("Zahlungsmethode")]
        public string PaymentMethod { get; set; } = string.Empty;

        [Name("Rechnungsbetrag (brutto)")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal InvoiceAmountGross { get; set; }

        [Name("Zahlbetrag")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal PaymentAmount { get; set; }

        [Name("Empfänger")]
        public string Recipient { get; set; } = string.Empty;

        [Name("Storno-Rechnungsnummer")]
        public string CancellationInvoiceNumber { get; set; } = string.Empty;

        [Name("Buchungsnr.")]
        public string BookingNumber { get; set; } = string.Empty;

        [Name("Forderungsbetrag")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal ClaimAmount { get; set; }

        [Name("davon beglichen")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal AmountPaid { get; set; }

        [Name("offener Betrag")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal OpenAmount { get; set; }

        [Name("Belegdatum")]
        [TypeConverter(typeof(GermanDateTimeConverter))]
        public DateTime DocumentDate { get; set; }

        [Name("Buchungsdatum")]
        [TypeConverter(typeof(GermanDateTimeConverter))]
        public DateTime BookingDate { get; set; }

        [Name("Vertrag")]
        public string Contract { get; set; } = string.Empty;

        [Name("Fälligkeitsdatum")]
        [TypeConverter(typeof(GermanDateTimeConverter))]
        public DateTime DueDate { get; set; }

        [Name("davon heute beglichen")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal AmountPaidToday { get; set; }

        [Name("davon beglichen am Stichtag")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal AmountPaidAtCutoffDate { get; set; }

        // Berechnete Eigenschaften für die UI
        public bool IsPaid => OpenAmount <= 0;
        public bool IsPartiallyPaid => AmountPaid > 0 && OpenAmount > 0;
        public bool IsCancelled => !string.IsNullOrEmpty(CancellationInvoiceNumber);
        public bool IsTokenPayment => PaymentMethod.Equals("Token", StringComparison.OrdinalIgnoreCase);

        // Für den Transaktionsvergleich (Email-Matching mit Stripe)
        public string CustomerEmail => Recipient;
        public decimal NetAmount => PaymentAmount;
        public string FormattedAmount => $"{NetAmount:C}";
        public string FormattedDate => DocumentDate.ToString("dd.MM.yyyy");

        // Status für UI-Anzeige
        public string PaymentStatus
        {
            get
            {
                if (IsCancelled) return "Storniert";
                if (IsPaid) return "Bezahlt";
                if (IsPartiallyPaid) return "Teilweise bezahlt";
                return "Offen";
            }
        }

        public string DisplayName => !string.IsNullOrEmpty(Contract) ? $"{Contract} ({Recipient})" : Recipient;

        public override string ToString()
        {
            return $"{Recipient} - {FormattedAmount} - {FormattedDate} - {PaymentStatus}";
        }
    }
}