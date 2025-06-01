using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using StripeCloud.Models;

namespace StripeCloud.Services
{
    public class ExcelExportService
    {
        /// <summary>
        /// Exportiert Vergleichsergebnisse in eine Excel-Datei
        /// </summary>
        public async Task<string> ExportComparisonResultsAsync(
            List<TransactionComparison> comparisons,
            ComparisonResult statistics,
            string fileName = "")
        {
            return await Task.Run(() =>
            {
                if (string.IsNullOrWhiteSpace(fileName))
                {
                    fileName = $"Transaktionsvergleich_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                }

                using var workbook = new XLWorkbook();

                // 1. Übersichts-Blatt
                CreateSummarySheet(workbook, statistics);

                // 2. Detailliertes Vergleichs-Blatt
                CreateComparisonSheet(workbook, comparisons);

                // 3. Probleme-Blatt (nur Diskrepanzen)
                CreateProblemsSheet(workbook, comparisons);

                // 4. Statistiken-Blatt
                CreateStatisticsSheet(workbook, comparisons, statistics);

                var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
                workbook.SaveAs(filePath);

                return filePath;
            });
        }

        /// <summary>
        /// Erstellt das Übersichts-Blatt
        /// </summary>
        private void CreateSummarySheet(XLWorkbook workbook, ComparisonResult statistics)
        {
            var worksheet = workbook.Worksheets.Add("Übersicht");

            // Header
            worksheet.Cell("A1").Value = "Transaktionsvergleich - Übersicht";
            worksheet.Cell("A1").Style.Font.FontSize = 16;
            worksheet.Cell("A1").Style.Font.Bold = true;

            worksheet.Cell("A3").Value = "Erstellt am:";
            worksheet.Cell("B3").Value = DateTime.Now.ToString("dd.MM.yyyy HH:mm");

            worksheet.Cell("A4").Value = "Verarbeitungsdauer:";
            worksheet.Cell("B4").Value = $"{statistics.ProcessingDuration.TotalSeconds:F1} Sekunden";

            // Statistiken
            worksheet.Cell("A6").Value = "Statistiken";
            worksheet.Cell("A6").Style.Font.Bold = true;

            worksheet.Cell("A8").Value = "Gesamt Transaktionen:";
            worksheet.Cell("B8").Value = statistics.TotalComparisons;

            worksheet.Cell("A9").Value = "Perfekte Übereinstimmungen:";
            worksheet.Cell("B9").Value = statistics.PerfectMatches;
            worksheet.Cell("C9").Value = $"{statistics.MatchRate:F1}%";

            worksheet.Cell("A10").Value = "Nur in Stripe:";
            worksheet.Cell("B10").Value = statistics.OnlyStripeCount;

            worksheet.Cell("A11").Value = "Nur in Chargecloud:";
            worksheet.Cell("B11").Value = statistics.OnlyChargecloudCount;

            worksheet.Cell("A12").Value = "Betragsabweichungen:";
            worksheet.Cell("B12").Value = statistics.AmountMismatchCount;

            // Finanzielle Übersicht
            worksheet.Cell("A14").Value = "Finanzielle Übersicht";
            worksheet.Cell("A14").Style.Font.Bold = true;

            worksheet.Cell("A16").Value = "Stripe Gesamtbetrag:";
            worksheet.Cell("B16").Value = statistics.TotalStripeAmount;
            worksheet.Cell("B16").Style.NumberFormat.Format = "#,##0.00 €";

            worksheet.Cell("A17").Value = "Chargecloud Gesamtbetrag:";
            worksheet.Cell("B17").Value = statistics.TotalChargecloudAmount;
            worksheet.Cell("B17").Style.NumberFormat.Format = "#,##0.00 €";

            worksheet.Cell("A18").Value = "Gesamtdiskrepanz:";
            worksheet.Cell("B18").Value = statistics.TotalDiscrepancy;
            worksheet.Cell("B18").Style.NumberFormat.Format = "#,##0.00 €";
            if (statistics.TotalDiscrepancy > 0)
            {
                worksheet.Cell("B18").Style.Font.FontColor = XLColor.Red;
            }

            // Formatierung
            worksheet.Columns().AdjustToContents();
        }

        /// <summary>
        /// Erstellt das detaillierte Vergleichs-Blatt
        /// </summary>
        private void CreateComparisonSheet(XLWorkbook workbook, List<TransactionComparison> comparisons)
        {
            var worksheet = workbook.Worksheets.Add("Vergleichsdetails");

            // Header
            var headers = new[]
            {
                "Kunde (E-Mail)", "Datum", "Betrag", "Status", "Stripe ✓", "Chargecloud ✓",
                "Betragsabweichung", "Stripe ID", "Stripe Betrag", "Stripe Datum",
                "Chargecloud Rechnung", "Chargecloud Betrag", "Chargecloud Datum", "Bemerkungen"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightBlue;
            }

            // Daten
            for (int i = 0; i < comparisons.Count; i++)
            {
                var comparison = comparisons[i];
                var row = i + 2;

                worksheet.Cell(row, 1).Value = comparison.CustomerEmail;
                worksheet.Cell(row, 2).Value = comparison.TransactionDate;
                worksheet.Cell(row, 3).Value = comparison.Amount;
                worksheet.Cell(row, 4).Value = comparison.StatusText;
                worksheet.Cell(row, 5).Value = comparison.StripeStatus;
                worksheet.Cell(row, 6).Value = comparison.ChargecloudStatus;
                worksheet.Cell(row, 7).Value = comparison.AmountDifference;
                worksheet.Cell(row, 8).Value = comparison.StripeTransaction?.Id ?? "";
                worksheet.Cell(row, 9).Value = comparison.StripeTransaction?.NetAmount ?? 0;
                worksheet.Cell(row, 10).Value = comparison.StripeTransaction?.CreatedDate;
                worksheet.Cell(row, 11).Value = comparison.ChargecloudTransaction?.InvoiceNumber ?? "";
                worksheet.Cell(row, 12).Value = comparison.ChargecloudTransaction?.NetAmount ?? 0;
                worksheet.Cell(row, 13).Value = comparison.ChargecloudTransaction?.DocumentDate;
                worksheet.Cell(row, 14).Value = GenerateNotes(comparison);

                // Farbkodierung basierend auf Status
                ApplyStatusFormatting(worksheet.Row(row), comparison.Status);
            }

            // Formatierung
            worksheet.Columns(2, 3).Style.DateFormat.Format = "dd.MM.yyyy";
            worksheet.Columns(10, 13).Style.DateFormat.Format = "dd.MM.yyyy";
            worksheet.Columns(3, 3).Style.NumberFormat.Format = "#,##0.00 €";
            worksheet.Columns(7, 7).Style.NumberFormat.Format = "#,##0.00 €";
            worksheet.Columns(9, 9).Style.NumberFormat.Format = "#,##0.00 €";
            worksheet.Columns(12, 12).Style.NumberFormat.Format = "#,##0.00 €";

            worksheet.Columns().AdjustToContents();
        }

        /// <summary>
        /// Erstellt das Probleme-Blatt (nur Diskrepanzen)
        /// </summary>
        private void CreateProblemsSheet(XLWorkbook workbook, List<TransactionComparison> comparisons)
        {
            var problems = comparisons.Where(c => c.Status != ComparisonStatus.Match).ToList();

            var worksheet = workbook.Worksheets.Add("Probleme");

            // Header
            worksheet.Cell("A1").Value = "Problematische Transaktionen";
            worksheet.Cell("A1").Style.Font.FontSize = 14;
            worksheet.Cell("A1").Style.Font.Bold = true;

            worksheet.Cell("A3").Value = $"Anzahl Probleme: {problems.Count}";

            // Gruppiert nach Problemtyp
            var groupedProblems = problems.GroupBy(p => p.Status).ToList();

            int currentRow = 5;

            foreach (var group in groupedProblems)
            {
                // Gruppen-Header
                worksheet.Cell(currentRow, 1).Value = GetStatusDisplayName(group.Key);
                worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                worksheet.Cell(currentRow, 1).Style.Fill.BackgroundColor = GetStatusColor(group.Key);
                currentRow += 2;

                // Headers für die Gruppe
                var headers = new[] { "Kunde", "Datum", "Betrag", "Details" };
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cell(currentRow, i + 1).Value = headers[i];
                    worksheet.Cell(currentRow, i + 1).Style.Font.Bold = true;
                }
                currentRow++;

                // Daten für die Gruppe
                foreach (var problem in group)
                {
                    worksheet.Cell(currentRow, 1).Value = problem.CustomerEmail;
                    worksheet.Cell(currentRow, 2).Value = problem.TransactionDate;
                    worksheet.Cell(currentRow, 3).Value = problem.Amount;
                    worksheet.Cell(currentRow, 4).Value = GenerateDetailedNotes(problem);
                    currentRow++;
                }

                currentRow += 2; // Abstand zur nächsten Gruppe
            }

            worksheet.Columns().AdjustToContents();
        }

        /// <summary>
        /// Erstellt das Statistiken-Blatt
        /// </summary>
        private void CreateStatisticsSheet(XLWorkbook workbook, List<TransactionComparison> comparisons, ComparisonResult statistics)
        {
            var worksheet = workbook.Worksheets.Add("Statistiken");

            // Kunden-Statistiken
            CreateCustomerStatistics(worksheet, comparisons);

            // Monatliche Übersicht
            CreateMonthlyStatistics(worksheet, comparisons);

            // Status-Verteilung
            CreateStatusDistribution(worksheet, comparisons);
        }

        private void CreateCustomerStatistics(IXLWorksheet worksheet, List<TransactionComparison> comparisons)
        {
            worksheet.Cell("A1").Value = "Kunden-Statistiken";
            worksheet.Cell("A1").Style.Font.Bold = true;

            var customerStats = comparisons
                .GroupBy(c => c.CustomerEmail)
                .Select(g => new
                {
                    Email = g.Key,
                    Total = g.Count(),
                    Matches = g.Count(c => c.Status == ComparisonStatus.Match),
                    Problems = g.Count(c => c.Status != ComparisonStatus.Match),
                    TotalAmount = g.Sum(c => Math.Abs(c.Amount))
                })
                .OrderByDescending(s => s.Problems)
                .ThenByDescending(s => s.TotalAmount)
                .ToList();

            // Headers
            var headers = new[] { "Kunde", "Gesamt", "Übereinstimmungen", "Probleme", "Gesamtbetrag" };
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(3, i + 1).Value = headers[i];
                worksheet.Cell(3, i + 1).Style.Font.Bold = true;
            }

            // Daten
            for (int i = 0; i < customerStats.Count; i++)
            {
                var stat = customerStats[i];
                var row = i + 4;

                worksheet.Cell(row, 1).Value = stat.Email;
                worksheet.Cell(row, 2).Value = stat.Total;
                worksheet.Cell(row, 3).Value = stat.Matches;
                worksheet.Cell(row, 4).Value = stat.Problems;
                worksheet.Cell(row, 5).Value = stat.TotalAmount;

                if (stat.Problems > 0)
                {
                    worksheet.Cell(row, 4).Style.Font.FontColor = XLColor.Red;
                }
            }

            worksheet.Columns().AdjustToContents();
        }

        private void CreateMonthlyStatistics(IXLWorksheet worksheet, List<TransactionComparison> comparisons)
        {
            worksheet.Cell("G1").Value = "Monatliche Übersicht";
            worksheet.Cell("G1").Style.Font.Bold = true;

            var monthlyStats = comparisons
                .GroupBy(c => new { c.TransactionDate.Year, c.TransactionDate.Month })
                .Select(g => new
                {
                    Period = $"{g.Key.Month:D2}/{g.Key.Year}",
                    Total = g.Count(),
                    Matches = g.Count(c => c.Status == ComparisonStatus.Match),
                    Amount = g.Sum(c => Math.Abs(c.Amount))
                })
                .OrderBy(s => s.Period)
                .ToList();

            // Headers
            var headers = new[] { "Monat", "Transaktionen", "Übereinstimmungen", "Betrag" };
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(3, i + 7).Value = headers[i];
                worksheet.Cell(3, i + 7).Style.Font.Bold = true;
            }

            // Daten
            for (int i = 0; i < monthlyStats.Count; i++)
            {
                var stat = monthlyStats[i];
                var row = i + 4;

                worksheet.Cell(row, 7).Value = stat.Period;
                worksheet.Cell(row, 8).Value = stat.Total;
                worksheet.Cell(row, 9).Value = stat.Matches;
                worksheet.Cell(row, 10).Value = stat.Amount;
            }
        }

        private void CreateStatusDistribution(IXLWorksheet worksheet, List<TransactionComparison> comparisons)
        {
            worksheet.Cell("A20").Value = "Status-Verteilung";
            worksheet.Cell("A20").Style.Font.Bold = true;

            var statusStats = comparisons
                .GroupBy(c => c.Status)
                .Select(g => new
                {
                    Status = GetStatusDisplayName(g.Key),
                    Count = g.Count(),
                    Percentage = (double)g.Count() / comparisons.Count * 100,
                    Amount = g.Sum(c => Math.Abs(c.Amount))
                })
                .OrderByDescending(s => s.Count)
                .ToList();

            // Headers
            var headers = new[] { "Status", "Anzahl", "Prozent", "Betrag" };
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(22, i + 1).Value = headers[i];
                worksheet.Cell(22, i + 1).Style.Font.Bold = true;
            }

            // Daten
            for (int i = 0; i < statusStats.Count; i++)
            {
                var stat = statusStats[i];
                var row = i + 23;

                worksheet.Cell(row, 1).Value = stat.Status;
                worksheet.Cell(row, 2).Value = stat.Count;
                worksheet.Cell(row, 3).Value = $"{stat.Percentage:F1}%";
                worksheet.Cell(row, 4).Value = stat.Amount;
            }
        }

        // Helper-Methoden
        private void ApplyStatusFormatting(IXLRow row, ComparisonStatus status)
        {
            var color = status switch
            {
                ComparisonStatus.Match => XLColor.LightGreen,
                ComparisonStatus.OnlyStripe => XLColor.LightYellow,
                ComparisonStatus.OnlyChargecloud => XLColor.LightYellow,
                ComparisonStatus.AmountMismatch => XLColor.LightPink,
                _ => XLColor.White
            };

            row.Style.Fill.BackgroundColor = color;
        }

        private XLColor GetStatusColor(ComparisonStatus status)
        {
            return status switch
            {
                ComparisonStatus.Match => XLColor.Green,
                ComparisonStatus.OnlyStripe => XLColor.Orange,
                ComparisonStatus.OnlyChargecloud => XLColor.Orange,
                ComparisonStatus.AmountMismatch => XLColor.Red,
                _ => XLColor.Gray
            };
        }

        private string GetStatusDisplayName(ComparisonStatus status)
        {
            return status switch
            {
                ComparisonStatus.Match => "Übereinstimmung",
                ComparisonStatus.OnlyStripe => "Nur in Stripe",
                ComparisonStatus.OnlyChargecloud => "Nur in Chargecloud",
                ComparisonStatus.AmountMismatch => "Betragsabweichung",
                _ => status.ToString()
            };
        }

        private string GenerateNotes(TransactionComparison comparison)
        {
            var notes = new List<string>();

            if (comparison.HasAmountDiscrepancy)
                notes.Add($"Differenz: {comparison.AmountDifference:C}");

            if (comparison.StripeTransaction?.IsRefunded == true)
                notes.Add("Stripe: Rückerstattet");

            if (comparison.ChargecloudTransaction?.IsCancelled == true)
                notes.Add("Chargecloud: Storniert");

            return string.Join("; ", notes);
        }

        private string GenerateDetailedNotes(TransactionComparison comparison)
        {
            return comparison.Status switch
            {
                ComparisonStatus.OnlyStripe => $"Stripe: {comparison.StripeTransaction?.Id} - {comparison.StripeTransaction?.FormattedAmount}",
                ComparisonStatus.OnlyChargecloud => $"Chargecloud: {comparison.ChargecloudTransaction?.InvoiceNumber} - {comparison.ChargecloudTransaction?.FormattedAmount}",
                ComparisonStatus.AmountMismatch => $"Stripe: {comparison.StripeTransaction?.FormattedAmount} vs Chargecloud: {comparison.ChargecloudTransaction?.FormattedAmount}",
                _ => ""
            };
        }
    }
}