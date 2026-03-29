using SentimentAnalyzerPro.Models;

namespace SentimentAnalyzerPro.Services
{
    public class SentimentService
    {
        private readonly MLService _mlService;
        private readonly TranslationService _translationService;

        public bool IsReady => _mlService.IsModelReady;

        public SentimentService()
        {
            _mlService = new MLService();
            _translationService = new TranslationService();
        }

        public async Task<string> InitializeAsync(IProgress<string>? progress = null)
        {
            return await _mlService.InitializeAsync(progress);
        }

        // Analyze single text
        public async Task<AnalysisResult> AnalyzeSingleAsync(string text)
        {
            var result = new AnalysisResult
            {
                OriginalText = text
            };

            try
            {
                // Detect and translate
                var (translatedText, detectedLang) = await _translationService.ProcessTextAsync(text);

                result.DetectedLanguage = detectedLang;
                result.TranslatedText = translatedText;

                // If non-english and no internet - skip
                if (detectedLang != "en" && translatedText == text && !TranslationService.IsInternetAvailable())
                {
                    result.Status = "Skipped";
                    result.StatusReason = "No internet for translation";
                    return result;
                }

                // ML prediction
                var prediction = _mlService.Predict(translatedText);
                if (prediction == null)
                {
                    result.Status = "Error";
                    result.StatusReason = "Model not ready";
                    return result;
                }

                result.IsPositive = prediction.Prediction;
                result.Confidence = Math.Max(prediction.Probability, 1 - prediction.Probability);
                result.Status = "Analyzed";
            }
            catch (Exception ex)
            {
                result.Status = "Error";
                result.StatusReason = ex.Message;
            }

            return result;
        }

        // Analyze bulk CSV
        public async Task<BulkAnalysisSummary> AnalyzeBulkAsync(
            string csvPath,
            IProgress<(int current, int total, string message)>? progress = null)
        {
            var summary = new BulkAnalysisSummary();

            // Read CSV
            var lines = await File.ReadAllLinesAsync(csvPath);
            // Skip header if exists
            var comments = lines
                .Skip(lines[0].Contains("text") || lines[0].Contains("review") || lines[0].Contains("comment") ? 1 : 0)
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .Select(l => l.Trim('"').Trim())
                .Where(l => l.Length > 2)
                .ToList();

            summary.TotalComments = comments.Count;

            for (int i = 0; i < comments.Count; i++)
            {
                progress?.Report((i + 1, comments.Count, $"Analyzing comment {i + 1} of {comments.Count}..."));

                var result = await AnalyzeSingleAsync(comments[i]);
                summary.Results.Add(result);

                if (result.Status == "Analyzed")
                {
                    if (result.IsPositive) summary.PositiveCount++;
                    else summary.NegativeCount++;
                }
                else
                {
                    summary.SkippedCount++;
                }
            }

            // Get top words
            var (posWord, negWord) = _mlService.GetTopWords(summary.Results);
            summary.TopPositiveWord = posWord;
            summary.TopNegativeWord = negWord;

            return summary;
        }

        // Export results to CSV
        public async Task ExportToCsvAsync(BulkAnalysisSummary summary, string outputPath)
        {
            var lines = new List<string>
            {
                "Original Text,Detected Language,Translated Text,Sentiment,Confidence,Status"
            };

            foreach (var r in summary.Results)
            {
                var lang = TranslationService.LanguageNames.TryGetValue(r.DetectedLanguage, out var name) ? name : r.DetectedLanguage;
                lines.Add($"\"{r.OriginalText.Replace("\"", "")}\",\"{lang}\",\"{r.TranslatedText.Replace("\"", "")}\",{r.SentimentLabel},{r.ConfidencePercent},{r.Status}");
            }

            await File.WriteAllLinesAsync(outputPath, lines);
        }
    }
}
