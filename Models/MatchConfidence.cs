using System;

namespace StripeCloud.Models
{
    /// <summary>
    /// Enum für die verschiedenen Match-Confidence Level
    /// </summary>
    public enum MatchConfidence
    {
        /// <summary>
        /// Layer 1: E-Mail + Betrag Match (99% sicher)
        /// </summary>
        High = 1,

        /// <summary>
        /// Layer 2: Name + Betrag Match (Name aus E-Mail abgeleitet)
        /// </summary>
        Medium = 2,

        /// <summary>
        /// Layer 3: Nur Betrag Match (niedrigste Sicherheit)
        /// </summary>
        Low = 3,

        /// <summary>
        /// Manuell bestätigter Match
        /// </summary>
        Manual = 4
    }

    /// <summary>
    /// Extension Methods für MatchConfidence
    /// </summary>
    public static class MatchConfidenceExtensions
    {
        public static string GetDisplayName(this MatchConfidence confidence)
        {
            return confidence switch
            {
                MatchConfidence.High => "✅ E-Mail + Betrag",
                MatchConfidence.Medium => "⚠️ Name + Betrag",
                MatchConfidence.Low => "⚠️ Nur Betrag",
                MatchConfidence.Manual => "✅ Manuell bestätigt",
                _ => "Unbekannt"
            };
        }

        public static string GetStatusColor(this MatchConfidence confidence)
        {
            return confidence switch
            {
                MatchConfidence.High => "#4CAF50",      // Satt-Grün
                MatchConfidence.Medium => "#8BC34A",    // Gelb-Grün  
                MatchConfidence.Low => "#FFC107",       // Helles Orange
                MatchConfidence.Manual => "#2E7D32",    // Dunkel-Grün
                _ => "#757575"
            };
        }

        public static string GetStatusText(this MatchConfidence confidence)
        {
            return confidence switch
            {
                MatchConfidence.High => "✅ Sichere Übereinstimmung",
                MatchConfidence.Medium => "⚠️ Name-Match (prüfen)",
                MatchConfidence.Low => "⚠️ Betrag-Match (prüfen)",
                MatchConfidence.Manual => "✅ Manuell bestätigt",
                _ => "Unbekannt"
            };
        }

        public static bool RequiresConfirmation(this MatchConfidence confidence)
        {
            return confidence == MatchConfidence.Medium || confidence == MatchConfidence.Low;
        }
    }
}