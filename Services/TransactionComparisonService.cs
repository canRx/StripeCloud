using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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
        /// AKTUALISIERT: Vergleicht Transaktionen für eine spezifische E-Mail-Adresse mit mehrstufigem Matching
        /// </summary>
        private List<TransactionComparison> CompareTransactionsForEmail(
            string email,
            List<StripeTransaction> stripeTransactions,
            List<ChargecloudTransaction> chargecloudTransactions)
        {
            var comparisons = new List<TransactionComparison>();
            var matchedStripe = new HashSet<StripeTransaction>();
            var matchedChargecloud = new HashSet<ChargecloudTransaction>();

            // LAYER 1: E-Mail + Betrag Matching (höchste Priorität)
            foreach (var stripe in stripeTransactions)
            {
                var exactMatch = FindExactEmailAndAmountMatch(stripe, chargecloudTransactions, matchedChargecloud);
                if (exactMatch != null)
                {
                    comparisons.Add(TransactionComparison.CreateMatched(stripe, exactMatch, MatchConfidence.High));
                    matchedStripe.Add(stripe);
                    matchedChargecloud.Add(exactMatch);
                }
            }

            // LAYER 2: Name + Betrag Matching (mittlere Priorität)
            foreach (var stripe in stripeTransactions.Where(s => !matchedStripe.Contains(s)))
            {
                var nameMatch = FindNameAndAmountMatch(stripe, chargecloudTransactions, matchedChargecloud);
                if (nameMatch != null)
                {
                    comparisons.Add(TransactionComparison.CreateMatched(stripe, nameMatch, MatchConfidence.Medium));
                    matchedStripe.Add(stripe);
                    matchedChargecloud.Add(nameMatch);
                }
            }

            // LAYER 3: Nur Betrag Matching (niedrigste Priorität)
            foreach (var stripe in stripeTransactions.Where(s => !matchedStripe.Contains(s)))
            {
                var amountMatch = FindAmountOnlyMatch(stripe, chargecloudTransactions, matchedChargecloud);
                if (amountMatch != null)
                {
                    comparisons.Add(TransactionComparison.CreateMatched(stripe, amountMatch, MatchConfidence.Low));
                    matchedStripe.Add(stripe);
                    matchedChargecloud.Add(amountMatch);
                }
            }

            // Nicht gematchte Stripe-Transaktionen hinzufügen
            foreach (var stripe in stripeTransactions.Where(s => !matchedStripe.Contains(s)))
            {
                comparisons.Add(TransactionComparison.CreateStripeOnly(stripe));
            }

            // Nicht gematchte Chargecloud-Transaktionen hinzufügen
            foreach (var chargecloud in chargecloudTransactions.Where(c => !matchedChargecloud.Contains(c)))
            {
                comparisons.Add(TransactionComparison.CreateChargecloudOnly(chargecloud));
            }

            return comparisons;
        }

        /// <summary>
        /// LAYER 1: Findet exakte E-Mail + Betrag Matches
        /// </summary>
        private ChargecloudTransaction? FindExactEmailAndAmountMatch(
            StripeTransaction stripe,
            List<ChargecloudTransaction> candidates,
            HashSet<ChargecloudTransaction> alreadyMatched)
        {
            var availableCandidates = candidates.Where(c => !alreadyMatched.Contains(c)).ToList();

            return availableCandidates
                .Where(c =>
                    string.Equals(stripe.CustomerEmail.Trim(), c.CustomerEmail.Trim(), StringComparison.OrdinalIgnoreCase) &&
                    Math.Abs(stripe.NetAmount - c.NetAmount) <= _amountMatchTolerance)
                .OrderBy(c => Math.Abs((c.DocumentDate - stripe.CreatedDate).TotalDays))
                .FirstOrDefault(c => Math.Abs((c.DocumentDate - stripe.CreatedDate).TotalDays) <= _dateMatchTolerance.TotalDays);
        }

        /// <summary>
        /// LAYER 2: Findet Name (aus E-Mail) + Betrag Matches
        /// </summary>
        private ChargecloudTransaction? FindNameAndAmountMatch(
            StripeTransaction stripe,
            List<ChargecloudTransaction> candidates,
            HashSet<ChargecloudTransaction> alreadyMatched)
        {
            var stripeName = ExtractNameFromEmail(stripe.CustomerEmail);
            if (string.IsNullOrEmpty(stripeName)) return null;

            var availableCandidates = candidates.Where(c => !alreadyMatched.Contains(c)).ToList();

            return availableCandidates
                .Where(c =>
                {
                    var chargecloudName = ExtractNameFromEmail(c.CustomerEmail);
                    return !string.IsNullOrEmpty(chargecloudName) &&
                           string.Equals(stripeName, chargecloudName, StringComparison.OrdinalIgnoreCase) &&
                           Math.Abs(stripe.NetAmount - c.NetAmount) <= _amountMatchTolerance;
                })
                .OrderBy(c => Math.Abs((c.DocumentDate - stripe.CreatedDate).TotalDays))
                .FirstOrDefault(c => Math.Abs((c.DocumentDate - stripe.CreatedDate).TotalDays) <= _dateMatchTolerance.TotalDays);
        }

        /// <summary>
        /// LAYER 3: Findet nur Betrag Matches
        /// </summary>
        private ChargecloudTransaction? FindAmountOnlyMatch(
            StripeTransaction stripe,
            List<ChargecloudTransaction> candidates,
            HashSet<ChargecloudTransaction> alreadyMatched)
        {
            var availableCandidates = candidates.Where(c => !alreadyMatched.Contains(c)).ToList();

            return availableCandidates
                .Where(c => Math.Abs(stripe.NetAmount - c.NetAmount) <= _amountMatchTolerance)
                .OrderBy(c => Math.Abs((c.DocumentDate - stripe.CreatedDate).TotalDays))
                .FirstOrDefault(c => Math.Abs((c.DocumentDate - stripe.CreatedDate).TotalDays) <= _dateMatchTolerance.TotalDays);
        }

        /// <summary>
        /// Extrahiert den Namen aus einer E-Mail-Adresse
        /// Beispiele: john.doe@example.com -> john.doe, j.smith@test.de -> j.smith
        /// </summary>
        private string ExtractNameFromEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email)) return string.Empty;

            try
            {
                var localPart = email.Split('@')[0];

                // Entferne Zahlen und Sonderzeichen, behalte nur Buchstaben, Punkte und Bindestriche
                var cleanName = Regex.Replace(localPart, @"[^a-zA-ZäöüÄÖÜß.\-]", "");

                // Wenn zu kurz oder nur Punkte/Bindestriche, return empty
                if (cleanName.Length < 2 || cleanName.Trim('.', '-').Length < 2)
                    return string.Empty;

                return cleanName.ToLowerInvariant();
            }
            catch
            {
                return string.Empty;
            }
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
                MatchConfidence = c.MatchConfidenceText, // NEU
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

            if (comparison.MatchConfidence.HasValue)
                notes.Add($"Match-Level: {comparison.MatchConfidence.Value.GetDisplayName()}");

            if (comparison.IsManuallyConfirmed)
                notes.Add("Manuell bestätigt");

            if (comparison.IsManuallyRejected)
                notes.Add("Manuell abgelehnt");

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

        // NEUE Statistiken für Match-Confidence
        public int HighConfidenceMatches { get; private set; }
        public int MediumConfidenceMatches { get; private set; }
        public int LowConfidenceMatches { get; private set; }
        public int PendingConfirmations { get; private set; }

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

            // NEUE Statistiken
            HighConfidenceMatches = Comparisons.Count(c => c.MatchConfidence == MatchConfidence.High);
            MediumConfidenceMatches = Comparisons.Count(c => c.MatchConfidence == MatchConfidence.Medium);
            LowConfidenceMatches = Comparisons.Count(c => c.MatchConfidence == MatchConfidence.Low);
            PendingConfirmations = Comparisons.Count(c => c.RequiresConfirmation);

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
                   $"Sicher: {HighConfidenceMatches}, " +
                   $"Zu prüfen: {PendingConfirmations}, " +
                   $"Nur Stripe: {OnlyStripeCount}, " +
                   $"Nur Chargecloud: {OnlyChargecloudCount}";
        }
    }

    /// <summary>
    /// AKTUALISIERTE Zeile für Excel-Export
    /// </summary>
    public class ExportRow
    {
        public string CustomerEmail { get; set; } = string.Empty;
        public DateTime TransactionDate { get; set; }
        public decimal Amount { get; set; }
        public string Status { get; set; } = string.Empty;
        public string MatchConfidence { get; set; } = string.Empty; // NEU
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