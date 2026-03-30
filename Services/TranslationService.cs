using Newtonsoft.Json;
using System.Net.NetworkInformation;
using System.Text;

namespace SentimentAnalyzerPro.Services
{
    public class TranslationService
    {
        private const string ApiUrl = "https://libretranslate.com/translate";
        private const string DetectUrl = "https://libretranslate.com/detect";
        private const string ApiKey = ""; // Optional

        private static readonly HttpClient _client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(8)
        };

        public static readonly Dictionary<string, string> LanguageNames = new()
        {
            { "en", "English" },
            { "ur", "Urdu" },
            { "hi", "Hindi" },
            { "ar", "Arabic" },
            { "fr", "French" },
            { "de", "German" },
            { "es", "Spanish" },
            { "zh", "Chinese" },
            { "pt", "Portuguese" },
            { "ru", "Russian" },
            { "ja", "Japanese" },
            { "ko", "Korean" },
            { "it", "Italian" },
            { "tr", "Turkish" },
            { "nl", "Dutch" },
            { "pl", "Polish" },
        };

        // Check internet
        public static bool IsInternetAvailable()
        {
            try
            {
                using var ping = new Ping();
                var reply = ping.Send("8.8.8.8", 2000);
                return reply.Status == IPStatus.Success;
            }
            catch { return false; }
        }

        // Detect language - first try script detection, then API
        public async Task<string> DetectLanguageAsync(string text)
        {
            // Step 1: Use script/character detection first (no API needed)
            string scriptDetected = DetectByScript(text);
            if (scriptDetected != "en" && scriptDetected != "unknown")
                return scriptDetected;

            // Step 2: If looks like English or unknown, try API
            if (!IsInternetAvailable()) return scriptDetected == "unknown" ? "en" : scriptDetected;

            try
            {
                var requestData = new { q = text };
                var json = JsonConvert.SerializeObject(requestData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                var response = await _client.PostAsync(DetectUrl, content);
                var responseText = await response.Content.ReadAsStringAsync();
                var results = JsonConvert.DeserializeObject<List<DetectionResult>>(responseText);
                if (results != null && results.Count > 0 && results[0].Confidence > 0.5f)
                    return results[0].Language ?? "en";
            }
            catch { }

            return "en";
        }

        // Detect language by Unicode script range - works 100% offline
        private string DetectByScript(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "en";

            int arabic = 0, urdu = 0, hindi = 0, chinese = 0,
                japanese = 0, korean = 0, russian = 0, latin = 0;

            foreach (char c in text)
            {
                // Arabic script (includes Urdu)
                if (c >= 0x0600 && c <= 0x06FF)
                {
                    // Urdu specific characters
                    if (c == 0x06AF || c == 0x06BA || c == 0x06BE ||
                        c == 0x06C1 || c == 0x06C3 || c == 0x06D2)
                        urdu++;
                    else
                        arabic++;
                }
                // Devanagari (Hindi)
                else if (c >= 0x0900 && c <= 0x097F) hindi++;
                // CJK Chinese/Japanese
                else if (c >= 0x4E00 && c <= 0x9FFF) chinese++;
                // Hiragana/Katakana (Japanese)
                else if ((c >= 0x3040 && c <= 0x309F) || (c >= 0x30A0 && c <= 0x30FF)) japanese++;
                // Korean Hangul
                else if (c >= 0xAC00 && c <= 0xD7AF) korean++;
                // Cyrillic (Russian)
                else if (c >= 0x0400 && c <= 0x04FF) russian++;
                // Latin
                else if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')) latin++;
            }

            int total = arabic + urdu + hindi + chinese + japanese + korean + russian + latin;
            if (total == 0) return "unknown";

            // Find dominant script
            if (urdu > total * 0.15) return "ur";
            if (arabic > total * 0.15) return "ar";
            if (hindi > total * 0.15) return "hi";
            if (japanese > total * 0.15) return "ja";
            if (chinese > total * 0.15) return "zh";
            if (korean > total * 0.15) return "ko";
            if (russian > total * 0.15) return "ru";

            // Latin - could be French, German, Spanish etc.
            // Use word-based detection for Latin scripts
            if (latin > total * 0.5)
                return DetectLatinLanguage(text.ToLower());

            return "unknown";
        }

        // Detect Latin-script languages by common words
        private string DetectLatinLanguage(string text)
        {
            var frenchWords = new[] { "le", "la", "les", "de", "du", "un", "une", "est", "et", "en", "je", "il", "elle", "nous", "vous", "ce", "que", "qui", "pas", "avec", "film", "tres", "pour" };
            var germanWords = new[] { "der", "die", "das", "und", "ist", "ein", "eine", "nicht", "ich", "sie", "wir", "mit", "auf", "von", "dem", "den", "des", "fur", "sehr", "aber" };
            var spanishWords = new[] { "el", "la", "los", "las", "de", "del", "en", "es", "un", "una", "que", "no", "con", "por", "para", "este", "esta", "muy", "bien", "producto" };
            var italianWords = new[] { "il", "la", "le", "di", "un", "una", "non", "con", "per", "che", "sono", "questo", "bella", "molto", "bello" };
            var portugueseWords = new[] { "de", "do", "da", "um", "uma", "nao", "com", "por", "para", "que", "este", "essa", "muito", "produto", "bom" };

            var words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);

            int fr = words.Count(w => frenchWords.Contains(w));
            int de = words.Count(w => germanWords.Contains(w));
            int es = words.Count(w => spanishWords.Contains(w));
            int it = words.Count(w => italianWords.Contains(w));
            int pt = words.Count(w => portugueseWords.Contains(w));

            int maxScore = Math.Max(fr, Math.Max(de, Math.Max(es, Math.Max(it, pt))));
            if (maxScore == 0) return "en";

            if (maxScore == fr) return "fr";
            if (maxScore == de) return "de";
            if (maxScore == es) return "es";
            if (maxScore == it) return "it";
            if (maxScore == pt) return "pt";

            return "en";
        }

        // Translate to English via API
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

                if (!response.IsSuccessStatusCode) return text;

                var responseText = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<TranslationResult>(responseText);
                return result?.TranslatedText ?? text;
            }
            catch { return text; }
        }

        // Main method
        public async Task<(string translatedText, string detectedLang)> ProcessTextAsync(string text)
        {
            string lang = await DetectLanguageAsync(text);

            if (lang == "en" || lang == "unknown")
                return (text, "en");

            // Non-English detected
            if (!IsInternetAvailable())
                return (text, lang); // Will be skipped by caller

            // Try translation
            string translated = await TranslateToEnglishAsync(text, lang);
            return (translated, lang);
        }

        private class DetectionResult
        {
            [JsonProperty("language")] public string? Language { get; set; }
            [JsonProperty("confidence")] public float Confidence { get; set; }
        }

        private class TranslationResult
        {
            [JsonProperty("translatedText")] public string? TranslatedText { get; set; }
        }
    }
}
