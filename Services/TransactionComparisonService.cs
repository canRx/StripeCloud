using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using StripeCloud.Models;

namespace StripeCloud.Services
{
    public class TransactionComparisonService
    {
        private readonly TimeSpan _dateMatchTolerance = TimeSpan.FromDays(7); // 7 Tage Toleranz
        private readonly decimal _amountMatchTolerance = 0.01m; // 1 Cent Toleranz

        /// <summary>
        /// Führt einen vollständigen Vergleich zwischen Stripe- und Chargecloud-Transaktionen durch
        /// </summary>
        public Task<ComparisonResult> CompareTransactionsAsync(
            List<StripeTransaction> stripeTransactions,
            List<ChargecloudTransaction> chargecloudTransactions)
        {
            var result = new ComparisonResult();
            var startTime = DateTime.Now;

            try
            {
                // 1. Nach E-Mail gruppieren
                var stripeByEmail = GroupStripeByEmail(stripeTransactions);
                var chargecloudByEmail = GroupChargecloudByEmail(chargecloudTransactions);

                // 2. Alle E-Mail-Adressen sammeln
                var allEmails = stripeByEmail.Keys.Union(chargecloudByEmail.Keys).ToList();

                // 3. Für jede E-Mail-Adresse Vergleich durchführen
                foreach (var email in allEmails)
                {
                    var stripeForEmail = stripeByEmail.GetValueOrDefault(email, new List<StripeTransaction>());
                    var chargecloudForEmail = chargecloudByEmail.GetValueOrDefault(email, new List<ChargecloudTransaction>());

                    var emailComparisons = CompareTransactionsForEmail(email, stripeForEmail, chargecloudForEmail);
                    result.Comparisons.AddRange(emailComparisons);
                }

                // 4. Statistiken berechnen
                result.CalculateStatistics();
                result.ProcessingDuration = DateTime.Now - startTime;

                return Task.FromResult(result);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Fehler beim Transaktionsvergleich: {ex.Message}");
                result.ProcessingDuration = DateTime.Now - startTime;
                return Task.FromResult(result);
            }
        }

        /// <summary>
        /// Vergleicht Transaktionen für eine spezifische E-Mail-Adresse
        /// </summary>
        /// <summary>
        /// Vergleicht Transaktionen für eine spezifische E-Mail-Adresse
        /// </summary>
        private List<TransactionComparison> CompareTransactionsForEmail(
            string email,
            List<StripeTransaction> stripeTransactions,
            List<ChargecloudTransaction> chargecloudTransactions)
        {
            var comparisons = new List<TransactionComparison>();
            var matchedStripe = new HashSet<StripeTransaction>();
            var matchedChargecloud = new HashSet<ChargecloudTransaction>();

            // 1. Versuche direkte Matches zu finden (Betrag + Datum)
            foreach (var stripe in stripeTransactions)
            {
                var matchingChargecloud = FindBestMatch(stripe, chargecloudTransactions, matchedChargecloud);

                if (matchingChargecloud != null)
                {
                    comparisons.Add(TransactionComparison.CreateMatched(stripe, matchingChargecloud));
                    matchedStripe.Add(stripe);
                    matchedChargecloud.Add(matchingChargecloud);
                }
            }

            // 2. Nicht gematchte Stripe-Transaktionen hinzufügen
            foreach (var stripe in stripeTransactions.Where(s => !matchedStripe.Contains(s)))
            {
                comparisons.Add(TransactionComparison.CreateStripeOnly(stripe));
            }

            // 3. Nicht gematchte Chargecloud-Transaktionen hinzufügen
            foreach (var chargecloud in chargecloudTransactions.Where(c => !matchedChargecloud.Contains(c)))
            {
                comparisons.Add(TransactionComparison.CreateChargecloudOnly(chargecloud));
            }

            return comparisons;
        }

        /// <summary>
        /// Findet die beste Übereinstimmung für eine Stripe-Transaktion
        /// </summary>
        private ChargecloudTransaction? FindBestMatch(
            StripeTransaction stripe,
            List<ChargecloudTransaction> candidates,
            HashSet<ChargecloudTransaction> alreadyMatched)
        {
            var availableCandidates = candidates.Where(c => !alreadyMatched.Contains(c)).ToList();

            if (!availableCandidates.Any())
                return null;

            // 1. Exakte Betrag-Matches bevorzugen
            var exactAmountMatches = availableCandidates
                .Where(c => Math.Abs(stripe.NetAmount - c.NetAmount) <= _amountMatchTolerance)
                .ToList();

            if (exactAmountMatches.Any())
            {
                // Von den exakten Betrag-Matches das mit dem nächsten Datum nehmen
                return exactAmountMatches
                    .OrderBy(c => Math.Abs((c.DocumentDate - stripe.CreatedDate).TotalDays))
                    .FirstOrDefault(c => Math.Abs((c.DocumentDate - stripe.CreatedDate).TotalDays) <= _dateMatchTolerance.TotalDays);
            }

            // 2. Wenn kein exakter Betrag-Match, dann beste Kombination aus Betrag und Datum
            var scoredMatches = availableCandidates
                .Select(c => new
                {
                    Transaction = c,
                    AmountDiff = Math.Abs(stripe.NetAmount - c.NetAmount),
                    DateDiff = Math.Abs((c.DocumentDate - stripe.CreatedDate).TotalDays),
                    Score = CalculateMatchScore(stripe, c)
                })
                .Where(m => m.DateDiff <= _dateMatchTolerance.TotalDays && m.Score > 0.5) // Mindest-Score für Match
                .OrderByDescending(m => m.Score)
                .ToList();

            return scoredMatches.FirstOrDefault()?.Transaction;
        }

        /// <summary>
        /// Berechnet einen Match-Score zwischen 0 und 1
        /// </summary>
        private double CalculateMatchScore(StripeTransaction stripe, ChargecloudTransaction chargecloud)
        {
            // Gewichtung: 60% Betrag, 40% Datum
            const double amountWeight = 0.6;
            const double dateWeight = 0.4;

            // Betrag-Score (je näher desto besser)
            var amountDiff = Math.Abs(stripe.NetAmount - chargecloud.NetAmount);
            var maxAmount = Math.Max(Math.Abs(stripe.NetAmount), Math.Abs(chargecloud.NetAmount));
            var amountScore = maxAmount > 0 ? Math.Max(0, 1 - (double)(amountDiff / maxAmount)) : 1;

            // Datum-Score (je näher desto besser)
            var dateDiff = Math.Abs((chargecloud.DocumentDate - stripe.CreatedDate).TotalDays);
            var dateScore = Math.Max(0, 1 - (dateDiff / _dateMatchTolerance.TotalDays));

            return (amountScore * amountWeight) + (dateScore * dateWeight);
        }

        /// <summary>
        /// Gruppiert Stripe-Transaktionen nach E-Mail
        /// </summary>
        private Dictionary<string, List<StripeTransaction>> GroupStripeByEmail(List<StripeTransaction> transactions)
        {
            return transactions
                .Where(t => !string.IsNullOrWhiteSpace(t.CustomerEmail))
                .GroupBy(t => t.CustomerEmail.ToLowerInvariant().Trim())
                .ToDictionary(g => g.Key, g => g.ToList());
        }

        /// <summary>
        /// Gruppiert Chargecloud-Transaktionen nach E-Mail
        /// </summary>
        private Dictionary<string, List<ChargecloudTransaction>> GroupChargecloudByEmail(List<ChargecloudTransaction> transactions)
        {
            return transactions
                .Where(t => !string.IsNullOrWhiteSpace(t.Recipient))
                .GroupBy(t => t.Recipient.ToLowerInvariant().Trim())
                .ToDictionary(g => g.Key, g => g.ToList());
        }

        /// <summary>
        /// Wendet Filter auf Vergleichsergebnisse an
        /// </summary>
        public List<TransactionComparison> ApplyFilters(List<TransactionComparison> comparisons, FilterOptions filters)
        {
            return comparisons.Where(filters.MatchesFilter).ToList();
        }

        /// <summary>
        /// Exportiert Vergleichsergebnisse für Excel/CSV
        /// </summary>
        public List<ExportRow> PrepareForExport(List<TransactionComparison> comparisons)
        {
            return comparisons.Select(c => new ExportRow
            {
                CustomerEmail = c.CustomerEmail,
                TransactionDate = c.TransactionDate,
                Amount = c.Amount,
                Status = c.StatusText,
                HasStripe = c.HasStripeTransaction,
                HasChargecloud = c.HasChargecloudTransaction,
                AmountDifference = c.AmountDifference,
                StripeId = c.StripeTransaction?.Id ?? "",
                StripeAmount = c.StripeTransaction?.NetAmount ?? 0,
                StripeDate = c.StripeTransaction?.CreatedDate,
                ChargecloudInvoice = c.ChargecloudTransaction?.InvoiceNumber ?? "",
                ChargecloudAmount = c.ChargecloudTransaction?.NetAmount ?? 0,
                ChargecloudDate = c.ChargecloudTransaction?.DocumentDate,
                Notes = GenerateNotes(c)
            }).ToList();
        }

        private string GenerateNotes(TransactionComparison comparison)
        {
            var notes = new List<string>();

            if (comparison.HasAmountDiscrepancy)
                notes.Add($"Betragsabweichung: {comparison.AmountDifference:C}");

            if (comparison.HasOnlyStripe)
                notes.Add("Nur in Stripe vorhanden");

            if (comparison.HasOnlyChargecloud)
                notes.Add("Nur in Chargecloud vorhanden");

            if (comparison.StripeTransaction?.IsRefunded == true)
                notes.Add($"Stripe: Rückerstattung {comparison.StripeTransaction.AmountRefunded:C}");

            return string.Join("; ", notes);
        }
    }

    /// <summary>
    /// Ergebnis des Transaktionsvergleichs
    /// </summary>
    public class ComparisonResult
    {
        public List<TransactionComparison> Comparisons { get; set; } = new();
        public List<string> Errors { get; set; } = new();
        public TimeSpan ProcessingDuration { get; set; }

        // Statistiken
        public int TotalComparisons => Comparisons.Count;
        public int PerfectMatches { get; private set; }
        public int OnlyStripeCount { get; private set; }
        public int OnlyChargecloudCount { get; private set; }
        public int AmountMismatchCount { get; private set; }
        public decimal TotalStripeAmount { get; private set; }
        public decimal TotalChargecloudAmount { get; private set; }
        public decimal TotalDiscrepancy { get; private set; }

        public bool IsSuccessful => Errors.Count == 0 && Comparisons.Count > 0;
        public double MatchRate => TotalComparisons > 0 ? (double)PerfectMatches / TotalComparisons * 100 : 0;

        internal void CalculateStatistics()
        {
            PerfectMatches = Comparisons.Count(c => c.Status == ComparisonStatus.Match);
            OnlyStripeCount = Comparisons.Count(c => c.Status == ComparisonStatus.OnlyStripe);
            OnlyChargecloudCount = Comparisons.Count(c => c.Status == ComparisonStatus.OnlyChargecloud);
            AmountMismatchCount = Comparisons.Count(c => c.Status == ComparisonStatus.AmountMismatch);

            TotalStripeAmount = Comparisons
                .Where(c => c.HasStripeTransaction)
                .Sum(c => c.StripeTransaction!.NetAmount);

            TotalChargecloudAmount = Comparisons
                .Where(c => c.HasChargecloudTransaction)
                .Sum(c => c.ChargecloudTransaction!.NetAmount);

            TotalDiscrepancy = Math.Abs(TotalStripeAmount - TotalChargecloudAmount);
        }

        public string GetSummary()
        {
            return $"Gesamt: {TotalComparisons}, " +
                   $"Übereinstimmungen: {PerfectMatches} ({MatchRate:F1}%), " +
                   $"Nur Stripe: {OnlyStripeCount}, " +
                   $"Nur Chargecloud: {OnlyChargecloudCount}, " +
                   $"Betragsabweichungen: {AmountMismatchCount}";
        }
    }

    /// <summary>
    /// Zeile für Excel-Export
    /// </summary>
    public class ExportRow
    {
        public string CustomerEmail { get; set; } = string.Empty;
        public DateTime TransactionDate { get; set; }
        public decimal Amount { get; set; }
        public string Status { get; set; } = string.Empty;
        public bool HasStripe { get; set; }
        public bool HasChargecloud { get; set; }
        public decimal AmountDifference { get; set; }
        public string StripeId { get; set; } = string.Empty;
        public decimal StripeAmount { get; set; }
        public DateTime? StripeDate { get; set; }
        public string ChargecloudInvoice { get; set; } = string.Empty;
        public decimal ChargecloudAmount { get; set; }
        public DateTime? ChargecloudDate { get; set; }
        public string Notes { get; set; } = string.Empty;
    }
}