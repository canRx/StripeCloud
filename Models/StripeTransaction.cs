using System;
using CsvHelper.Configuration.Attributes;
using StripeCloud.Helpers;

namespace StripeCloud.Models
{
    public class StripeTransaction
    {
        [Name("id")]
        public string Id { get; set; } = string.Empty;

        [Name("Created date (UTC)")]
        [TypeConverter(typeof(GermanDateTimeConverter))]
        public DateTime CreatedDate { get; set; }

        [Name("Amount")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal Amount { get; set; }

        [Name("Amount Refunded")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal AmountRefunded { get; set; }

        [Name("Currency")]
        public string Currency { get; set; } = string.Empty;

        [Name("Captured")]
        [TypeConverter(typeof(SafeBooleanConverter))]
        public bool Captured { get; set; }

        [Name("Converted Amount")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal ConvertedAmount { get; set; }

        [Name("Converted Amount Refunded")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal ConvertedAmountRefunded { get; set; }

        [Name("Converted Currency")]
        public string ConvertedCurrency { get; set; } = string.Empty;

        [Name("Decline Reason")]
        public string DeclineReason { get; set; } = string.Empty;

        [Name("Description")]
        public string Description { get; set; } = string.Empty;

        [Name("Fee")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal Fee { get; set; }

        [Name("Refunded date (UTC)")]
        [TypeConverter(typeof(SafeDateTimeConverter))]
        public DateTime? RefundedDate { get; set; }

        [Name("Statement Descriptor")]
        public string StatementDescriptor { get; set; } = string.Empty;

        [Name("Status")]
        public string Status { get; set; } = string.Empty;

        [Name("Seller Message")]
        public string SellerMessage { get; set; } = string.Empty;

        [Name("Taxes On Fee")]
        [TypeConverter(typeof(GermanDecimalConverter))]
        public decimal TaxesOnFee { get; set; }

        [Name("Card ID")]
        public string CardId { get; set; } = string.Empty;

        [Name("Customer ID")]
        public string CustomerId { get; set; } = string.Empty;

        [Name("Customer Description")]
        public string CustomerDescription { get; set; } = string.Empty;

        [Name("Customer Email")]
        public string CustomerEmail { get; set; } = string.Empty;

        [Name("Invoice ID")]
        public string InvoiceId { get; set; } = string.Empty;

        [Name("Transfer")]
        public string Transfer { get; set; } = string.Empty;

        // Berechnete Eigenschaften für die UI
        public decimal NetAmount => Amount - AmountRefunded;
        public bool IsRefunded => AmountRefunded > 0;
        public bool IsSuccessful => Status.Equals("Paid", StringComparison.OrdinalIgnoreCase);

        // Für den Transaktionsvergleich
        public string DisplayName => !string.IsNullOrEmpty(CustomerDescription) ? CustomerDescription : CustomerEmail;
        public string FormattedAmount => $"{NetAmount:C} {Currency?.ToUpper()}";
        public string FormattedDate => CreatedDate.ToString("dd.MM.yyyy HH:mm");

        public override string ToString()
        {
            return $"{CustomerEmail} - {FormattedAmount} - {FormattedDate}";
        }
    }
}