using System;
using System.Globalization;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;

namespace StripeCloud.Helpers
{
    public class GermanDecimalConverter : ITypeConverter
    {
        public object ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0m;

            // Entferne Anführungszeichen falls vorhanden
            text = text.Trim('"').Trim();

            // KORRIGIERT: Bessere Erkennung des Chargecloud-Formats
            if (IsChargecloudFormat(text))
            {
                return ParseChargecloudFormat(text);
            }

            // Deutsches Format prüfen: "80,00" oder "3.688,50"
            if (IsGermanFormat(text))
            {
                return ParseGermanFormat(text);
            }

            // Standard englische Formatierung: "80.00" oder "3,688.50"
            if (decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal englishResult))
                return englishResult;

            // Fallback: Punkt durch Komma ersetzen und als deutsch versuchen
            string fallbackText = text.Replace('.', ',');
            if (decimal.TryParse(fallbackText, NumberStyles.Number, new CultureInfo("de-DE"), out decimal fallbackResult))
                return fallbackResult;

            // Wenn alles fehlschlägt, gib 0 zurück
            return 0m;
        }

        /// <summary>
        /// Erkennt das Chargecloud-Format: genau 6 Nachkommastellen ohne Tausendertrennzeichen
        /// Beispiele: "17.370000", "123.456000", "0.120000"
        /// </summary>
        private bool IsChargecloudFormat(string text)
        {
            if (!text.Contains('.'))
                return false;

            var parts = text.Split('.');

            // Muss genau 2 Teile haben (Vorkomma.Nachkomma)
            if (parts.Length != 2)
                return false;

            var wholePart = parts[0];
            var decimalPart = parts[1];

            // Nachkommateil muss genau 6 Stellen haben
            if (decimalPart.Length != 6)
                return false;

            // Vorkommateil darf keine Leerzeichen oder weitere Punkte enthalten
            // und sollte eine vernünftige Länge haben (max 10 Stellen für normale Beträge)
            if (wholePart.Length > 10 || !IsNumeric(wholePart) || !IsNumeric(decimalPart))
                return false;

            // Zusätzliche Validierung: Wenn Vorkommateil mehr als 3 Stellen hat,
            // könnte es deutsches Format mit Tausendertrennzeichen sein
            // Beispiel: "3.688" sollte nicht als Chargecloud erkannt werden
            if (wholePart.Length > 3 && !text.EndsWith("000000"))
            {
                // Wenn es nicht mit 000000 endet, ist es wahrscheinlich deutsches Format
                return false;
            }

            return true;
        }

        /// <summary>
        /// Erkennt deutsches Zahlenformat mit Komma als Dezimaltrennzeichen
        /// Beispiele: "80,00", "3.688,50", "1.234.567,89"
        /// </summary>
        private bool IsGermanFormat(string text)
        {
            // Enthält Komma als mögliches Dezimaltrennzeichen
            if (!text.Contains(','))
                return false;

            // Komma sollte nicht am Ende oder Anfang stehen
            if (text.StartsWith(',') || text.EndsWith(','))
                return false;

            // Splitten am Komma - sollte genau 2 Teile geben
            var parts = text.Split(',');
            if (parts.Length != 2)
                return false;

            var wholePart = parts[0];
            var decimalPart = parts[1];

            // Nachkommateil sollte 1-3 Stellen haben (typisch 2)
            if (decimalPart.Length < 1 || decimalPart.Length > 3)
                return false;

            // Beide Teile müssen numerisch sein (Punkte als Tausendertrennzeichen erlaubt)
            if (!IsNumericWithThousandsSeparator(wholePart) || !IsNumeric(decimalPart))
                return false;

            return true;
        }

        /// <summary>
        /// Parst Chargecloud-Format: "17.370000" -> 17.37
        /// </summary>
        private decimal ParseChargecloudFormat(string text)
        {
            var parts = text.Split('.');
            var wholePart = parts[0];
            var decimalPart = parts[1].Substring(0, 2); // Nur die ersten 2 Stellen

            var correctedText = $"{wholePart}.{decimalPart}";

            if (decimal.TryParse(correctedText, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return 0m;
        }

        /// <summary>
        /// Parst deutsches Format: "3.688,50" -> 3688.50
        /// </summary>
        private decimal ParseGermanFormat(string text)
        {
            if (decimal.TryParse(text, NumberStyles.Number, new CultureInfo("de-DE"), out decimal result))
                return result;

            return 0m;
        }

        /// <summary>
        /// Prüft ob ein String nur Ziffern enthält
        /// </summary>
        private bool IsNumeric(string text)
        {
            return !string.IsNullOrEmpty(text) && text.All(char.IsDigit);
        }

        /// <summary>
        /// Prüft ob ein String numerisch ist und Punkte als Tausendertrennzeichen haben kann
        /// Beispiele: "123", "1.234", "1.234.567"
        /// </summary>
        private bool IsNumericWithThousandsSeparator(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            // Entferne alle Punkte und prüfe ob der Rest numerisch ist
            var withoutDots = text.Replace(".", "");
            if (!IsNumeric(withoutDots))
                return false;

            // Prüfe Tausendertrennzeichen-Regeln
            if (text.Contains('.'))
            {
                var parts = text.Split('.');

                // Erstes Teil kann 1-3 Stellen haben
                if (parts[0].Length < 1 || parts[0].Length > 3)
                    return false;

                // Alle weiteren Teile müssen genau 3 Stellen haben (außer dem letzten)
                for (int i = 1; i < parts.Length; i++)
                {
                    if (parts[i].Length != 3)
                        return false;
                }
            }

            return true;
        }

        public string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData)
        {
            if (value is decimal decimalValue)
                return decimalValue.ToString("F2", CultureInfo.InvariantCulture);

            return string.Empty;
        }
    }

    public class GermanDateTimeConverter : ITypeConverter
    {
        public object ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData)
        {
            if (string.IsNullOrWhiteSpace(text))
                return DateTime.MinValue;

            // Entferne Anführungszeichen falls vorhanden
            text = text.Trim('"');

            // Standard ISO Format
            if (DateTime.TryParse(text, out DateTime result))
                return result;

            return DateTime.MinValue;
        }

        public string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData)
        {
            if (value is DateTime dateTime)
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss");

            return string.Empty;
        }
    }

    public class SafeBooleanConverter : ITypeConverter
    {
        public object ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            // Entferne Anführungszeichen falls vorhanden
            text = text.Trim('"').ToLowerInvariant();

            // Verschiedene Boolean-Formate unterstützen
            return text switch
            {
                "true" => true,
                "false" => false,
                "1" => true,
                "0" => false,
                "yes" => true,
                "no" => false,
                "ja" => true,
                "nein" => false,
                _ => false
            };
        }

        public string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData)
        {
            if (value is bool boolValue)
                return boolValue.ToString().ToLowerInvariant();

            return "false";
        }
    }

    public class SafeDateTimeConverter : ITypeConverter
    {
        public object? ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            // Entferne Anführungszeichen falls vorhanden
            text = text.Trim('"');

            // Standard ISO Format
            if (DateTime.TryParse(text, out DateTime result))
                return result;

            return null;
        }

        public string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData)
        {
            if (value is DateTime dateTime)
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss");

            return string.Empty;
        }
    }
}