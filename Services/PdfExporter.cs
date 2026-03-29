using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Colors;
using SentimentAnalyzerPro.Models;

namespace SentimentAnalyzerPro.Services
{
    public static class PdfExporter
    {
        public static void Export(BulkAnalysisSummary summary, string outputPath)
        {
            using var writer = new PdfWriter(outputPath);
            using var pdf = new PdfDocument(writer);
            using var doc = new Document(pdf);

            var titleColor = new DeviceRgb(99, 102, 241);
            var greenColor = new DeviceRgb(34, 197, 94);
            var redColor = new DeviceRgb(239, 68, 68);
            var grayColor = new DeviceRgb(100, 100, 120);
            var altBg = new DeviceRgb(245, 245, 255);

            // Title
            doc.Add(new Paragraph("Sentiment Analysis Report")
                .SetFontSize(22).SetFontColor(titleColor).SetBold()
                .SetTextAlignment(TextAlignment.CENTER));

            doc.Add(new Paragraph($"Generated: {DateTime.Now:dd MMM yyyy, HH:mm}")
                .SetFontSize(10).SetFontColor(grayColor)
                .SetTextAlignment(TextAlignment.CENTER));

            doc.Add(new Paragraph("\n"));

            // Summary
            doc.Add(new Paragraph("Summary").SetFontSize(14).SetBold().SetFontColor(titleColor));

            var summaryTable = new Table(2).SetWidth(UnitValue.CreatePercentValue(60));
            AddSummaryRow(summaryTable, "Total Comments", summary.TotalComments.ToString());
            AddSummaryRow(summaryTable, "Positive", $"{summary.PositiveCount} ({summary.PositivePercent:F1}%)");
            AddSummaryRow(summaryTable, "Negative", $"{summary.NegativeCount} ({summary.NegativePercent:F1}%)");
            AddSummaryRow(summaryTable, "Skipped", summary.SkippedCount.ToString());
            AddSummaryRow(summaryTable, "Overall Mood", summary.OverallMood);
            doc.Add(summaryTable);
            doc.Add(new Paragraph("\n"));

            // Results
            doc.Add(new Paragraph("Detailed Results").SetFontSize(14).SetBold().SetFontColor(titleColor));

            var table = new Table(new float[] { 4, 1.5f, 1.5f, 1.5f })
                .SetWidth(UnitValue.CreatePercentValue(100));

            foreach (var header in new[] { "Comment", "Language", "Sentiment", "Confidence" })
            {
                table.AddHeaderCell(new Cell()
                    .Add(new Paragraph(header).SetBold().SetFontColor(ColorConstants.WHITE))
                    .SetBackgroundColor(titleColor).SetPadding(6));
            }

            bool alt = false;
            foreach (var r in summary.Results.Take(200))
            {
                var bg = alt ? altBg : ColorConstants.WHITE;
                string shortText = r.OriginalText.Length > 70 ? r.OriginalText[..70] + "..." : r.OriginalText;
                string langName = TranslationService.LanguageNames.TryGetValue(r.DetectedLanguage, out var n) ? n : r.DetectedLanguage;

                table.AddCell(new Cell().Add(new Paragraph(shortText).SetFontSize(9)).SetBackgroundColor(bg).SetPadding(5));
                table.AddCell(new Cell().Add(new Paragraph(langName).SetFontSize(9)).SetBackgroundColor(bg).SetPadding(5));
                table.AddCell(new Cell()
                    .Add(new Paragraph(r.SentimentLabel).SetFontSize(9)
                    .SetFontColor(r.IsPositive ? greenColor : redColor).SetBold())
                    .SetBackgroundColor(bg).SetPadding(5));
                table.AddCell(new Cell().Add(new Paragraph(r.ConfidencePercent).SetFontSize(9)).SetBackgroundColor(bg).SetPadding(5));
                alt = !alt;
            }

            doc.Add(table);

            if (summary.Results.Count > 200)
                doc.Add(new Paragraph($"\n(Showing first 200 of {summary.Results.Count} results)")
                    .SetFontSize(9).SetFontColor(grayColor));
        }

        private static void AddSummaryRow(Table table, string key, string value)
        {
            table.AddCell(new Cell().Add(new Paragraph(key).SetBold().SetFontSize(10)).SetPadding(5));
            table.AddCell(new Cell().Add(new Paragraph(value).SetFontSize(10)).SetPadding(5));
        }
    }
}
