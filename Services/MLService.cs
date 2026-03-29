using Microsoft.ML;
using Microsoft.ML.Data;
using SentimentAnalyzerPro.Models;

namespace SentimentAnalyzerPro.Services
{
    public class MLService
    {
        private readonly MLContext _mlContext;
        private ITransformer? _model;
        private PredictionEngine<ReviewData, SentimentPrediction>? _predictor;

        private static readonly string ModelPath = "sentiment_model.zip";
        private static readonly string DatasetPath = Path.Combine("Data", "IMDB Dataset.csv");

        public bool IsModelReady => _model != null && _predictor != null;

        public MLService()
        {
            _mlContext = new MLContext(seed: 42);
        }

        public async Task<string> InitializeAsync(IProgress<string>? progress = null)
        {
            return await Task.Run(() =>
            {
                try
                {
                    if (File.Exists(ModelPath))
                    {
                        progress?.Report("Loading existing model...");
                        LoadModel();
                        return "Model loaded successfully!";
                    }

                    if (!File.Exists(DatasetPath))
                        return $"Dataset not found at: {DatasetPath}";

                    progress?.Report("Reading and cleaning dataset...");
                    var records = ParseImdbCsv(DatasetPath, progress);

                    progress?.Report($"Loaded {records.Count} reviews. Building pipeline...");
                    var dataView = _mlContext.Data.LoadFromEnumerable(records);

                    progress?.Report("Splitting data...");
                    var split = _mlContext.Data.TrainTestSplit(dataView, testFraction: 0.2);

                    progress?.Report("Building ML pipeline...");
                    var pipeline = _mlContext.Transforms.Text
                        .FeaturizeText("Features", nameof(ReviewData.Text))
                        .Append(_mlContext.BinaryClassification.Trainers
                            .SdcaLogisticRegression(
                                labelColumnName: "Label",
                                featureColumnName: "Features"));

                    progress?.Report("Training model (1-2 minutes)...");
                    _model = pipeline.Fit(split.TrainSet);

                    progress?.Report("Evaluating accuracy...");
                    var predictions = _model.Transform(split.TestSet);
                    var metrics = _mlContext.BinaryClassification.Evaluate(predictions, "Label");

                    progress?.Report($"Accuracy: {metrics.Accuracy:P2} — Saving model...");
                    _mlContext.Model.Save(_model, dataView.Schema, ModelPath);

                    _predictor = _mlContext.Model
                        .CreatePredictionEngine<ReviewData, SentimentPrediction>(_model);

                    return $"Model trained! Accuracy: {metrics.Accuracy:P2}";
                }
                catch (Exception ex)
                {
                    return $"Error: {ex.Message}";
                }
            });
        }

        private List<ReviewData> ParseImdbCsv(string path, IProgress<string>? progress)
        {
            var records = new List<ReviewData>();
            var allText = File.ReadAllText(path);
            var lines = allText.Split('\n');

            bool firstLine = true;
            foreach (var line in lines)
            {
                if (firstLine) { firstLine = false; continue; }

                var trimmed = line.Trim().Trim('\r');
                if (string.IsNullOrWhiteSpace(trimmed)) continue;

                int lastComma = trimmed.LastIndexOf(',');
                if (lastComma < 0) continue;

                string sentiment = trimmed.Substring(lastComma + 1).Trim().Trim('"').ToLower();
                string reviewText = trimmed.Substring(0, lastComma).Trim().Trim('"');

                reviewText = System.Text.RegularExpressions.Regex.Replace(reviewText, "<.*?>", " ");
                reviewText = reviewText.Replace("&quot;", "\"").Replace("&#39;", "'").Replace("&amp;", "&");

                if (string.IsNullOrWhiteSpace(reviewText)) continue;
                if (sentiment != "positive" && sentiment != "negative") continue;

                records.Add(new ReviewData
                {
                    Text = reviewText,
                    Sentiment = sentiment == "positive"
                });

                if (records.Count % 5000 == 0)
                    progress?.Report($"Parsed {records.Count} reviews...");

                if (records.Count >= 40000) break;
            }

            return records;
        }

        private void LoadModel()
        {
            _model = _mlContext.Model.Load(ModelPath, out _);
            _predictor = _mlContext.Model
                .CreatePredictionEngine<ReviewData, SentimentPrediction>(_model);
        }

        public SentimentPrediction? Predict(string text)
        {
            if (_predictor == null) return null;
            var input = new ReviewData { Text = text };
            return _predictor.Predict(input);
        }

        public (string positive, string negative) GetTopWords(List<AnalysisResult> results)
        {
            var positiveWords = results
                .Where(r => r.IsPositive && r.Status == "Analyzed")
                .SelectMany(r => r.TranslatedText.ToLower().Split(' '))
                .Where(w => w.Length > 4)
                .GroupBy(w => w)
                .OrderByDescending(g => g.Count())
                .FirstOrDefault()?.Key ?? "N/A";

            var negativeWords = results
                .Where(r => !r.IsPositive && r.Status == "Analyzed")
                .SelectMany(r => r.TranslatedText.ToLower().Split(' '))
                .Where(w => w.Length > 4)
                .GroupBy(w => w)
                .OrderByDescending(g => g.Count())
                .FirstOrDefault()?.Key ?? "N/A";

            return (positiveWords, negativeWords);
        }
    }
}