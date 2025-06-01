using System;
using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;

namespace StripeCloud.Helpers
{
    public class GermanDecimalConverter : ITypeConverter
    {
        public object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0m;

            // Entferne Anführungszeichen falls vorhanden
            text = text.Trim('"').Trim();

            // Spezialbehandlung für Chargecloud-Format: 17.370000 -> 17.37
            if (text.Contains('.') && text.Length > 6)
            {
                // Prüfe ob es das Chargecloud-Format ist (Punkt + 6 Stellen)
                var parts = text.Split('.');
                if (parts.Length == 2 && parts[1].Length == 6)
                {
                    // Das ist Chargecloud-Format: 17.370000 bedeutet 17.37
                    var wholePart = parts[0];
                    var decimalPart = parts[1].Substring(0, 2); // Nur die ersten 2 Stellen
                    var correctedText = $"{wholePart}.{decimalPart}";

                    if (decimal.TryParse(correctedText, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal chargecloudResult))
                        return chargecloudResult;
                }
            }

            // Stripe-Format: "80,00" (deutsche Formatierung)
            if (decimal.TryParse(text, NumberStyles.Number, new CultureInfo("de-DE"), out decimal result))
                return result;

            // Standard englische Formatierung
            if (decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out result))
                return result;

            // Punkt durch Komma ersetzen und nochmal versuchen
            string germanText = text.Replace('.', ',');
            if (decimal.TryParse(germanText, NumberStyles.Number, new CultureInfo("de-DE"), out result))
                return result;

            // Wenn alles fehlschlägt, gib 0 zurück
            return 0m;
        }

        public string ConvertToString(object value, IWriterRow row, MemberMapData memberMapData)
        {
            if (value is decimal decimalValue)
                return decimalValue.ToString("F2", CultureInfo.InvariantCulture);

            return string.Empty;
        }
    }

    public class GermanDateTimeConverter : ITypeConverter
    {
        public object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
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

        public string ConvertToString(object value, IWriterRow row, MemberMapData memberMapData)
        {
            if (value is DateTime dateTime)
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss");

            return string.Empty;
        }
    }

    public class SafeBooleanConverter : ITypeConverter
    {
        public object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
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

        public string ConvertToString(object value, IWriterRow row, MemberMapData memberMapData)
        {
            if (value is bool boolValue)
                return boolValue.ToString().ToLowerInvariant();

            return "false";
        }
    }

    public class SafeDateTimeConverter : ITypeConverter
    {
        public object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
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

        public string ConvertToString(object value, IWriterRow row, MemberMapData memberMapData)
        {
            if (value is DateTime dateTime)
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss");

            return string.Empty;
        }
    }
}