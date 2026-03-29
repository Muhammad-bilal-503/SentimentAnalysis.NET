using Newtonsoft.Json;
using System.Net.NetworkInformation;
using System.Text;

namespace SentimentAnalyzerPro.Services
{
    public class TranslationService
    {
        // LibreTranslate free public API
        private const string ApiUrl = "https://libretranslate.com/translate";
        private const string DetectUrl = "https://libretranslate.com/detect";
        private const string ApiKey = ""; // Leave empty for free tier

        private static readonly HttpClient _client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(10)
        };

        // Language codes to display names
        public static readonly Dictionary<string, string> LanguageNames = new()
        {
            { "en", "English 🇬🇧" },
            { "ur", "Urdu 🇵🇰" },
            { "hi", "Hindi 🇮🇳" },
            { "ar", "Arabic 🇸🇦" },
            { "fr", "French 🇫🇷" },
            { "de", "German 🇩🇪" },
            { "es", "Spanish 🇪🇸" },
            { "zh", "Chinese 🇨🇳" },
            { "pt", "Portuguese 🇧🇷" },
            { "ru", "Russian 🇷🇺" },
            { "ja", "Japanese 🇯🇵" },
            { "ko", "Korean 🇰🇷" },
        };

        // Check internet connection
        public static bool IsInternetAvailable()
        {
            try
            {
                using var ping = new Ping();
                var reply = ping.Send("8.8.8.8", 2000);
                return reply.Status == IPStatus.Success;
            }
            catch
            {
                return false;
            }
        }

        // Detect language of text
        public async Task<string> DetectLanguageAsync(string text)
        {
            try
            {
                var requestData = new { q = text };
                var json = JsonConvert.SerializeObject(requestData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await _client.PostAsync(DetectUrl, content);
                var responseText = await response.Content.ReadAsStringAsync();

                var results = JsonConvert.DeserializeObject<List<DetectionResult>>(responseText);
                if (results != null && results.Count > 0)
                    return results[0].Language ?? "en";

                return "en";
            }
            catch
            {
                return "en"; // Default to English on error
            }
        }

        // Translate text to English
        public async Task<string> TranslateToEnglishAsync(string text, string sourceLang)
        {
            try
            {
                var requestData = new
                {
                    q = text,
                    source = sourceLang,
                    target = "en",
                    format = "text",
                    api_key = ApiKey
                };

                var json = JsonConvert.SerializeObject(requestData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await _client.PostAsync(ApiUrl, content);
                var responseText = await response.Content.ReadAsStringAsync();

                var result = JsonConvert.DeserializeObject<TranslationResult>(responseText);
                return result?.TranslatedText ?? text;
            }
            catch
            {
                return text; // Return original on error
            }
        }

        // Main method: detect + translate if needed
        public async Task<(string translatedText, string detectedLang)> ProcessTextAsync(string text)
        {
            // Detect language first
            string lang = await DetectLanguageAsync(text);

            // Already English - no translation needed
            if (lang == "en")
                return (text, "en");

            // Not English - check internet
            if (!IsInternetAvailable())
                return (text, lang); // Return original, let caller handle skip

            // Translate to English
            string translated = await TranslateToEnglishAsync(text, lang);
            return (translated, lang);
        }

        // Helper classes for JSON deserialization
        private class DetectionResult
        {
            [JsonProperty("language")]
            public string? Language { get; set; }

            [JsonProperty("confidence")]
            public float Confidence { get; set; }
        }

        private class TranslationResult
        {
            [JsonProperty("translatedText")]
            public string? TranslatedText { get; set; }
        }
    }
}
