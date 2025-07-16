using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using StripeCloud.Models;

namespace StripeCloud.Services
{
    public class CsvImportService
    {
        public async Task<List<StripeTransaction>> ImportStripeTransactionsAsync(string filePath)
        {
            try
            {
                // Erst Format validieren
                if (!await IsValidStripeFormatAsync(filePath))
                {
                    throw new InvalidDataException("Die Datei entspricht nicht dem erwarteten Stripe-Format.");
                }

                using var reader = new StreamReader(filePath);
                using var csv = new CsvReader(reader, GetStripeConfiguration());

                var transactions = new List<StripeTransaction>();

                await foreach (var record in csv.GetRecordsAsync<StripeTransaction>())
                {
                    // Validierung und Säuberung der Daten
                    if (IsValidStripeTransaction(record))
                    {
                        transactions.Add(record);
                    }
                }

                return transactions.OrderByDescending(t => t.CreatedDate).ToList();
            }
            catch (InvalidDataException)
            {
                // User-freundliche Fehlermeldung weiterleiten
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Unbekanntes CSV-Format - Die ausgewählte Datei kann nicht gelesen werden.", ex);
            }
        }

        public async Task<List<ChargecloudTransaction>> ImportChargecloudTransactionsAsync(string filePath)
        {
            try
            {
                // Erst Format validieren
                if (!await IsValidChargecloudFormatAsync(filePath))
                {
                    throw new InvalidDataException("Die Datei entspricht nicht dem erwarteten Chargecloud-Format.");
                }

                using var reader = new StreamReader(filePath);
                using var csv = new CsvReader(reader, GetChargecloudConfiguration());

                var transactions = new List<ChargecloudTransaction>();

                await foreach (var record in csv.GetRecordsAsync<ChargecloudTransaction>())
                {
                    // Validierung und Säuberung der Daten
                    if (IsValidChargecloudTransaction(record))
                    {
                        transactions.Add(record);
                    }
                }

                return transactions.OrderByDescending(t => t.DocumentDate).ToList();
            }
            catch (InvalidDataException)
            {
                // User-freundliche Fehlermeldung weiterleiten
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Unbekanntes CSV-Format - Die ausgewählte Datei kann nicht gelesen werden.", ex);
            }
        }

        // Format-Validierung
        private async Task<bool> IsValidStripeFormatAsync(string filePath)
        {
            try
            {
                using var reader = new StreamReader(filePath);
                var firstLine = await reader.ReadLineAsync();

                if (string.IsNullOrWhiteSpace(firstLine))
                    return false;

                // Prüfe auf wichtige Stripe-Header
                var requiredHeaders = new[] {
                    "Created date (UTC)",
                    "Customer Email",
                    "Amount",
                    "Status"
                };

                return requiredHeaders.All(header =>
                    firstLine.Contains(header, StringComparison.OrdinalIgnoreCase));
            }
            catch
            {
                return false;
            }
        }

        private async Task<bool> IsValidChargecloudFormatAsync(string filePath)
        {
            try
            {
                using var reader = new StreamReader(filePath);
                var firstLine = await reader.ReadLineAsync();

                if (string.IsNullOrWhiteSpace(firstLine))
                    return false;

                // Prüfe auf wichtige Chargecloud-Header
                var requiredHeaders = new[] {
                    "Rechnungsnummer",
                    "Empfänger",
                    "Zahlungsmethode",
                    "Belegdatum"
                };

                return requiredHeaders.All(header =>
                    firstLine.Contains(header, StringComparison.OrdinalIgnoreCase));
            }
            catch
            {
                return false;
            }
        }

        private static CsvConfiguration GetStripeConfiguration()
        {
            return new CsvConfiguration(new CultureInfo("de-DE"))
            {
                Delimiter = ",",
                HasHeaderRecord = true,
                IgnoreBlankLines = true,
                TrimOptions = TrimOptions.Trim,
                MissingFieldFound = null,
                BadDataFound = null,
                Quote = '"'
            };
        }

        private static CsvConfiguration GetChargecloudConfiguration()
        {
            return new CsvConfiguration(new CultureInfo("de-DE"))
            {
                Delimiter = ";",
                HasHeaderRecord = true,
                IgnoreBlankLines = true,
                TrimOptions = TrimOptions.Trim,
                MissingFieldFound = null,
                BadDataFound = null
            };
        }

        private static bool IsValidStripeTransaction(StripeTransaction transaction)
        {
            // Basis-Validierung
            if (string.IsNullOrWhiteSpace(transaction.Id))
                return false;

            if (string.IsNullOrWhiteSpace(transaction.CustomerEmail))
                return false;

            // Nur erfolgreiche oder teilweise erfolgreiche Transaktionen
            if (!IsStripeTransactionRelevant(transaction))
                return false;

            return true;
        }

        private static bool IsValidChargecloudTransaction(ChargecloudTransaction transaction)
        {
            // Basis-Validierung
            if (string.IsNullOrWhiteSpace(transaction.InvoiceNumber))
                return false;

            if (string.IsNullOrWhiteSpace(transaction.Recipient))
                return false;

            // Nur Token-Zahlungen sind relevant für Stripe-Vergleich
            if (!transaction.IsTokenPayment)
                return false;

            // Keine stornierten Rechnungen
            if (transaction.IsCancelled)
                return false;

            return true;
        }

        private static bool IsStripeTransactionRelevant(StripeTransaction transaction)
        {
            // Erfolgreiche Zahlungen
            if (transaction.Status.Equals("succeeded", StringComparison.OrdinalIgnoreCase) ||
                transaction.Status.Equals("paid", StringComparison.OrdinalIgnoreCase))
                return true;

            // Erfasste Zahlungen
            if (transaction.Captured && transaction.Amount > 0)
                return true;

            // Teilweise zurückerstattete Zahlungen (noch relevanter Betrag vorhanden)
            if (transaction.NetAmount > 0)
                return true;

            return false;
        }

        // Automatische Erkennung des CSV-Formats
        public async Task<CsvFileType> DetectCsvTypeAsync(string filePath)
        {
            try
            {
                if (await IsValidStripeFormatAsync(filePath))
                    return CsvFileType.Stripe;

                if (await IsValidChargecloudFormatAsync(filePath))
                    return CsvFileType.Chargecloud;

                return CsvFileType.Unknown;
            }
            catch
            {
                return CsvFileType.Unknown;
            }
        }

        // Vorschau der ersten paar Zeilen für Benutzer-Bestätigung
        public async Task<List<string>> GetCsvPreviewAsync(string filePath, int maxLines = 5)
        {
            var preview = new List<string>();

            try
            {
                using var reader = new StreamReader(filePath);
                for (int i = 0; i < maxLines && !reader.EndOfStream; i++)
                {
                    var line = await reader.ReadLineAsync();
                    if (line != null)
                        preview.Add(line);
                }
            }
            catch (Exception ex)
            {
                preview.Add($"Fehler beim Lesen der Datei: {ex.Message}");
            }

            return preview;
        }

        // Import-Statistiken mit besserem Error Handling
        public class ImportResult<T> where T : class
        {
            public List<T> Data { get; set; } = new();
            public int TotalRows { get; set; }
            public int SuccessfulRows { get; set; }
            public int SkippedRows { get; set; }
            public List<string> Errors { get; set; } = new();
            public TimeSpan ImportDuration { get; set; }

            public bool IsSuccessful => Data.Count > 0 && !HasCriticalErrors;
            public bool HasCriticalErrors => Errors.Any(e => e.Contains("Unbekanntes CSV-Format") || e.Contains("entspricht nicht dem erwarteten"));
            public double SuccessRate => TotalRows > 0 ? (double)SuccessfulRows / TotalRows * 100 : 0;
        }

        public async Task<ImportResult<StripeTransaction>> ImportStripeWithStatsAsync(string filePath)
        {
            var startTime = DateTime.Now;
            var result = new ImportResult<StripeTransaction>();

            try
            {
                // Format-Validierung zuerst
                if (!await IsValidStripeFormatAsync(filePath))
                {
                    result.Errors.Add("❌ Unbekanntes CSV-Format\n\nDie ausgewählte Datei kann nicht gelesen werden.\nBitte stellen Sie sicher, dass es sich um eine\ngültige Stripe-CSV-Datei handelt.");
                    result.ImportDuration = DateTime.Now - startTime;
                    return result;
                }

                using var reader = new StreamReader(filePath);
                using var csv = new CsvReader(reader, GetStripeConfiguration());

                await foreach (var record in csv.GetRecordsAsync<StripeTransaction>())
                {
                    result.TotalRows++;

                    try
                    {
                        if (IsValidStripeTransaction(record))
                        {
                            result.Data.Add(record);
                            result.SuccessfulRows++;
                        }
                        else
                        {
                            result.SkippedRows++;
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Errors.Add($"Zeile {result.TotalRows}: {ex.Message}");
                        result.SkippedRows++;
                    }
                }

                result.Data = result.Data.OrderByDescending(t => t.CreatedDate).ToList();
            }
            catch (Exception)
            {
                // KORRIGIERT: Exception-Variable entfernt, da sie nicht verwendet wird
                result.Errors.Add("❌ Unbekanntes CSV-Format\n\nDie ausgewählte Datei kann nicht gelesen werden.\nBitte stellen Sie sicher, dass es sich um eine\ngültige Stripe-CSV-Datei handelt.");
            }

            result.ImportDuration = DateTime.Now - startTime;
            return result;
        }

        public async Task<ImportResult<ChargecloudTransaction>> ImportChargecloudWithStatsAsync(string filePath)
        {
            var startTime = DateTime.Now;
            var result = new ImportResult<ChargecloudTransaction>();

            try
            {
                // Format-Validierung zuerst
                if (!await IsValidChargecloudFormatAsync(filePath))
                {
                    result.Errors.Add("❌ Unbekanntes CSV-Format\n\nDie ausgewählte Datei kann nicht gelesen werden.\nBitte stellen Sie sicher, dass es sich um eine\ngültige Chargecloud-CSV-Datei handelt.");
                    result.ImportDuration = DateTime.Now - startTime;
                    return result;
                }

                using var reader = new StreamReader(filePath);
                using var csv = new CsvReader(reader, GetChargecloudConfiguration());

                await foreach (var record in csv.GetRecordsAsync<ChargecloudTransaction>())
                {
                    result.TotalRows++;

                    try
                    {
                        if (IsValidChargecloudTransaction(record))
                        {
                            result.Data.Add(record);
                            result.SuccessfulRows++;
                        }
                        else
                        {
                            result.SkippedRows++;
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Errors.Add($"Zeile {result.TotalRows}: {ex.Message}");
                        result.SkippedRows++;
                    }
                }

                result.Data = result.Data.OrderByDescending(t => t.DocumentDate).ToList();
            }
            catch (Exception)
            {
                // KORRIGIERT: Exception-Variable entfernt, da sie nicht verwendet wird
                result.Errors.Add("❌ Unbekanntes CSV-Format\n\nDie ausgewählte Datei kann nicht gelesen werden.\nBitte stellen Sie sicher, dass es sich um eine\ngültige Chargecloud-CSV-Datei handelt.");
            }

            result.ImportDuration = DateTime.Now - startTime;
            return result;
        }
    }

    public enum CsvFileType
    {
        Unknown,
        Stripe,
        Chargecloud
    }
}