using System;
using System.Collections.Generic;
using System.Linq;
using StripeCloud.Models;

namespace StripeCloud.Services
{
    public class ManualMatchingService
    {
        /// <summary>
        /// Validiert und führt ein manuelles Matching von Transaktionen durch
        /// </summary>
        public ManualMatchingResult MatchTransactions(
            List<TransactionComparison> allComparisons,
            List<TransactionComparison> selectedTransactions)
        {
            var result = new ManualMatchingResult();

            try
            {
                // 1. Validierung der Auswahl
                var validationResult = ValidateSelection(selectedTransactions);
                if (!validationResult.IsValid)
                {
                    result.IsSuccessful = false;
                    result.ErrorMessage = validationResult.ErrorMessage;
                    return result;
                }

                // 2. Extrahiere Stripe und Chargecloud Transaktionen
                var stripeTransactions = selectedTransactions
                    .Where(t => t.HasStripeTransaction)
                    .Select(t => t.StripeTransaction!)
                    .ToList();

                var chargecloudTransactions = selectedTransactions
                    .Where(t => t.HasChargecloudTransaction)
                    .Select(t => t.ChargecloudTransaction!)
                    .ToList();

                // 3. Erstelle neue gematchte Transaktionen
                var newMatches = CreateManualMatches(stripeTransactions, chargecloudTransactions);

                // 4. Entferne alte Transaktionen und füge neue hinzu
                var updatedComparisons = UpdateComparisons(allComparisons, selectedTransactions, newMatches);

                result.IsSuccessful = true;
                result.UpdatedComparisons = updatedComparisons;
                result.MatchedCount = newMatches.Count;
                result.MatchedTransactions = newMatches;

                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccessful = false;
                result.ErrorMessage = $"Fehler beim manuellen Matching: {ex.Message}";
                return result;
            }
        }

        /// <summary>
        /// Validiert die Auswahl für manuelles Matching
        /// </summary>
        private ValidationResult ValidateSelection(List<TransactionComparison> selectedTransactions)
        {
            var result = new ValidationResult();

            // Mindestens 2 Transaktionen müssen ausgewählt sein
            if (selectedTransactions.Count < 2)
            {
                result.IsValid = false;
                result.ErrorMessage = "Mindestens 2 Transaktionen müssen ausgewählt werden.";
                return result;
            }

            // Maximal 10 Transaktionen (Performance-Schutz)
            if (selectedTransactions.Count > 10)
            {
                result.IsValid = false;
                result.ErrorMessage = "Maximal 10 Transaktionen können gleichzeitig gematcht werden.";
                return result;
            }

            // Zähle Stripe und Chargecloud Transaktionen
            var stripeCount = selectedTransactions.Count(t => t.HasStripeTransaction);
            var chargecloudCount = selectedTransactions.Count(t => t.HasChargecloudTransaction);

            // Es müssen sowohl Stripe als auch Chargecloud Transaktionen vorhanden sein
            if (stripeCount == 0)
            {
                result.IsValid = false;
                result.ErrorMessage = "Mindestens eine Stripe-Transaktion muss ausgewählt werden.";
                return result;
            }

            if (chargecloudCount == 0)
            {
                result.IsValid = false;
                result.ErrorMessage = "Mindestens eine Chargecloud-Transaktion muss ausgewählt werden.";
                return result;
            }

            // Warnungen für ungleiche Anzahl
            if (stripeCount != chargecloudCount)
            {
                result.HasWarnings = true;
                result.WarningMessage = $"Ungleiche Anzahl: {stripeCount} Stripe und {chargecloudCount} Chargecloud Transaktionen. " +
                                      "Einige Transaktionen bleiben unverknüpft.";
            }

            // Prüfe auf bereits gematchte Transaktionen
            var alreadyMatched = selectedTransactions.Where(t => t.Status == ComparisonStatus.Match).ToList();
            if (alreadyMatched.Any())
            {
                result.HasWarnings = true;
                result.WarningMessage = $"{alreadyMatched.Count} bereits gematchte Transaktionen werden neu verknüpft.";
            }

            // Prüfe auf doppelte Transaktionen
            var duplicateStripe = selectedTransactions
                .Where(t => t.HasStripeTransaction)
                .GroupBy(t => t.StripeTransaction!.Id)
                .Where(g => g.Count() > 1)
                .ToList();

            var duplicateChargecloud = selectedTransactions
                .Where(t => t.HasChargecloudTransaction)
                .GroupBy(t => t.ChargecloudTransaction!.InvoiceNumber)
                .Where(g => g.Count() > 1)
                .ToList();

            if (duplicateStripe.Any() || duplicateChargecloud.Any())
            {
                result.IsValid = false;
                result.ErrorMessage = "Doppelte Transaktionen in der Auswahl gefunden.";
                return result;
            }

            result.IsValid = true;
            return result;
        }

        /// <summary>
        /// Erstellt manuelle Matches aus den ausgewählten Transaktionen
        /// </summary>
        private List<TransactionComparison> CreateManualMatches(
            List<StripeTransaction> stripeTransactions,
            List<ChargecloudTransaction> chargecloudTransactions)
        {
            var matches = new List<TransactionComparison>();

            // Sortiere beide Listen nach Datum für bessere Zuordnung
            var sortedStripe = stripeTransactions.OrderBy(t => t.CreatedDate).ToList();
            var sortedChargecloud = chargecloudTransactions.OrderBy(t => t.DocumentDate).ToList();

            // Erstelle Matches (1:1 Zuordnung)
            var maxMatches = Math.Min(sortedStripe.Count, sortedChargecloud.Count);

            for (int i = 0; i < maxMatches; i++)
            {
                var match = TransactionComparison.CreateMatched(
                    sortedStripe[i],
                    sortedChargecloud[i],
                    MatchConfidence.Manual);

                matches.Add(match);
            }

            // Übrige Stripe-Transaktionen (falls mehr Stripe als Chargecloud)
            for (int i = maxMatches; i < sortedStripe.Count; i++)
            {
                matches.Add(TransactionComparison.CreateStripeOnly(sortedStripe[i]));
            }

            // Übrige Chargecloud-Transaktionen (falls mehr Chargecloud als Stripe)
            for (int i = maxMatches; i < sortedChargecloud.Count; i++)
            {
                matches.Add(TransactionComparison.CreateChargecloudOnly(sortedChargecloud[i]));
            }

            return matches;
        }

        /// <summary>
        /// Aktualisiert die gesamte Comparison-Liste mit den neuen Matches
        /// </summary>
        private List<TransactionComparison> UpdateComparisons(
            List<TransactionComparison> allComparisons,
            List<TransactionComparison> selectedTransactions,
            List<TransactionComparison> newMatches)
        {
            var updatedComparisons = new List<TransactionComparison>();

            // Füge alle nicht-ausgewählten Transaktionen hinzu
            foreach (var comparison in allComparisons)
            {
                if (!selectedTransactions.Contains(comparison))
                {
                    updatedComparisons.Add(comparison);
                }
            }

            // Füge die neuen Matches hinzu
            updatedComparisons.AddRange(newMatches);

            return updatedComparisons;
        }

        /// <summary>
        /// Hebt ein manuelles Match auf und erstellt separate Transaktionen
        /// </summary>
        public ManualMatchingResult UnmatchTransactions(
            List<TransactionComparison> allComparisons,
            List<TransactionComparison> selectedMatches)
        {
            var result = new ManualMatchingResult();

            try
            {
                // Validierung: Nur bereits gematchte Transaktionen können aufgehoben werden
                var invalidSelections = selectedMatches
                    .Where(t => t.Status != ComparisonStatus.Match || !t.HasBothTransactions)
                    .ToList();

                if (invalidSelections.Any())
                {
                    result.IsSuccessful = false;
                    result.ErrorMessage = "Nur vollständig gematchte Transaktionen können aufgehoben werden.";
                    return result;
                }

                var newTransactions = new List<TransactionComparison>();

                // Erstelle separate Transaktionen für jedes aufgehobene Match
                foreach (var match in selectedMatches)
                {
                    if (match.StripeTransaction != null)
                    {
                        newTransactions.Add(TransactionComparison.CreateStripeOnly(match.StripeTransaction));
                    }

                    if (match.ChargecloudTransaction != null)
                    {
                        newTransactions.Add(TransactionComparison.CreateChargecloudOnly(match.ChargecloudTransaction));
                    }
                }

                // Aktualisiere die Comparison-Liste
                var updatedComparisons = new List<TransactionComparison>();

                // Füge alle nicht-ausgewählten Transaktionen hinzu
                foreach (var comparison in allComparisons)
                {
                    if (!selectedMatches.Contains(comparison))
                    {
                        updatedComparisons.Add(comparison);
                    }
                }

                // Füge die neuen separaten Transaktionen hinzu
                updatedComparisons.AddRange(newTransactions);

                result.IsSuccessful = true;
                result.UpdatedComparisons = updatedComparisons;
                result.UnmatchedCount = selectedMatches.Count;
                result.MatchedTransactions = newTransactions;

                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccessful = false;
                result.ErrorMessage = $"Fehler beim Aufheben des Matches: {ex.Message}";
                return result;
            }
        }
    }

    /// <summary>
    /// Ergebnis der manuellen Matching-Operation
    /// </summary>
    public class ManualMatchingResult
    {
        public bool IsSuccessful { get; set; } = false;
        public string ErrorMessage { get; set; } = string.Empty;
        public List<TransactionComparison> UpdatedComparisons { get; set; } = new();
        public List<TransactionComparison> MatchedTransactions { get; set; } = new();
        public int MatchedCount { get; set; } = 0;
        public int UnmatchedCount { get; set; } = 0;
    }

    /// <summary>
    /// Validierungsergebnis für die Auswahl
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; set; } = true;
        public string ErrorMessage { get; set; } = string.Empty;
        public bool HasWarnings { get; set; } = false;
        public string WarningMessage { get; set; } = string.Empty;
    }
}