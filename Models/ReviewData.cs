using Microsoft.ML.Data;

namespace SentimentAnalyzerPro.Models
{
    // Input data class - matches CSV columns
    public class ReviewData
    {
        [LoadColumn(0)]
        public string Text { get; set; } = string.Empty;

        [LoadColumn(1), ColumnName("Label")]
        public bool Sentiment { get; set; }
    }

    // Output prediction class
    public class SentimentPrediction : ReviewData
    {
        [ColumnName("PredictedLabel")]
        public bool Prediction { get; set; }

        public float Probability { get; set; }

        public float Score { get; set; }
    }

    // Result for each comment analyzed
    public class AnalysisResult
    {
        public string OriginalText { get; set; } = string.Empty;
        public string TranslatedText { get; set; } = string.Empty;
        public string DetectedLanguage { get; set; } = "en";
        public bool IsPositive { get; set; }
        public float Confidence { get; set; }
        public string Status { get; set; } = "Analyzed"; // "Analyzed", "Skipped", "Error"
        public string StatusReason { get; set; } = string.Empty;

        public string SentimentLabel => IsPositive ? "Positive" : "Negative";
        public string ConfidencePercent => $"{Confidence * 100:F1}%";
    }

    // Summary stats for bulk analysis
    public class BulkAnalysisSummary
    {
        public int TotalComments { get; set; }
        public int PositiveCount { get; set; }
        public int NegativeCount { get; set; }
        public int SkippedCount { get; set; }
        public int AnalyzedCount => PositiveCount + NegativeCount;
        public double PositivePercent => AnalyzedCount > 0 ? (double)PositiveCount / AnalyzedCount * 100 : 0;
        public double NegativePercent => AnalyzedCount > 0 ? (double)NegativeCount / AnalyzedCount * 100 : 0;
        public string TopPositiveWord { get; set; } = string.Empty;
        public string TopNegativeWord { get; set; } = string.Empty;
        public string OverallMood => PositivePercent >= 50 ? "😊 POSITIVE" : "😞 NEGATIVE";
        public List<AnalysisResult> Results { get; set; } = new();
    }
}
