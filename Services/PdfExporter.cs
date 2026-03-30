using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using ITextColor = iText.Kernel.Colors.Color;
using SentimentAnalyzerPro.Models;

namespace SentimentAnalyzerPro.Services
{
    public static class PdfExporter
    {
        public static void Export(BulkAnalysisSummary summary, string outputPath)
        {
            try
            {
                var writerProps = new WriterProperties();
                using var writer = new PdfWriter(outputPath, writerProps);
                using var pdf = new PdfDocument(writer);
                using var doc = new Document(pdf);

                // Fonts
                var titleFont  = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                var boldFont   = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                var normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

                // Colors — all typed as DeviceRgb, no ambiguity
                DeviceRgb titleColor  = new DeviceRgb(99, 102, 241);
                DeviceRgb greenColor  = new DeviceRgb(34, 139, 34);
                DeviceRgb redColor    = new DeviceRgb(200, 50, 50);
                DeviceRgb grayColor   = new DeviceRgb(100, 100, 120);
                DeviceRgb headerBg    = new DeviceRgb(99, 102, 241);
                DeviceRgb altBg       = new DeviceRgb(240, 240, 255);
                DeviceRgb lightBg     = new DeviceRgb(255, 255, 255);

                // ===== TITLE =====
                doc.Add(new Paragraph("Sentiment Analysis Report")
                    .SetFont(titleFont)
                    .SetFontSize(22)
                    .SetFontColor(titleColor)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetMarginBottom(5));

                doc.Add(new Paragraph($"Generated: {DateTime.Now:dd MMM yyyy, HH:mm}")
                    .SetFont(normalFont)
                    .SetFontSize(10)
                    .SetFontColor(grayColor)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetMarginBottom(20));

                // ===== SUMMARY =====
                doc.Add(new Paragraph("Summary")
                    .SetFont(boldFont)
                    .SetFontSize(14)
                    .SetFontColor(titleColor)
                    .SetMarginBottom(8));

                var summaryTable = new Table(
                    UnitValue.CreatePercentArray(new float[] { 50, 50 }))
                    .SetWidth(UnitValue.CreatePercentValue(55))
                    .SetMarginBottom(20);

                AddSummaryRow(summaryTable, "Total Comments",
                    summary.TotalComments.ToString(), boldFont, normalFont);
                AddSummaryRow(summaryTable, "Positive",
                    $"{summary.PositiveCount} ({summary.PositivePercent:F1}%)",
                    boldFont, normalFont);
                AddSummaryRow(summaryTable, "Negative",
                    $"{summary.NegativeCount} ({summary.NegativePercent:F1}%)",
                    boldFont, normalFont);
                AddSummaryRow(summaryTable, "Skipped",
                    summary.SkippedCount.ToString(), boldFont, normalFont);
                AddSummaryRow(summaryTable, "Overall Mood",
                    summary.OverallMood
                        .Replace("😊", "").Replace("😞", "").Trim(),
                    boldFont, normalFont);
                AddSummaryRow(summaryTable, "Top Positive Word",
                    summary.TopPositiveWord, boldFont, normalFont);
                AddSummaryRow(summaryTable, "Top Negative Word",
                    summary.TopNegativeWord, boldFont, normalFont);

                doc.Add(summaryTable);

                // ===== DETAILED RESULTS =====
                doc.Add(new Paragraph("Detailed Results")
                    .SetFont(boldFont)
                    .SetFontSize(14)
                    .SetFontColor(titleColor)
                    .SetMarginBottom(8));

                var table = new Table(
                    UnitValue.CreatePercentArray(
                        new float[] { 40, 15, 15, 15, 15 }))
                    .SetWidth(UnitValue.CreatePercentValue(100))
                    .SetMarginBottom(10);

                // Header row
                foreach (var header in new[]
                    { "Comment", "Language", "Sentiment", "Confidence", "Status" })
                {
                    table.AddHeaderCell(new Cell()
                        .Add(new Paragraph(header)
                            .SetFont(boldFont)
                            .SetFontSize(9)
                            .SetFontColor(ColorConstants.WHITE))
                        .SetBackgroundColor(headerBg)
                        .SetPadding(5));
                }

                // Data rows
                bool alt = false;
                foreach (var r in summary.Results.Take(300))
                {
                    DeviceRgb rowBg = alt ? altBg : lightBg;

                    string shortText = r.OriginalText.Length > 60
                        ? r.OriginalText[..60] + "..."
                        : r.OriginalText;
                    shortText = CleanForPdf(shortText);

                    string langName = TranslationService.LanguageNames
                        .TryGetValue(r.DetectedLanguage, out var n)
                        ? n : r.DetectedLanguage;

                    DeviceRgb sentColor = r.IsPositive ? greenColor : redColor;

                    table.AddCell(MakeCell(shortText,   normalFont, 8, rowBg));
                    table.AddCell(MakeCell(langName,    normalFont, 8, rowBg));

                    // Sentiment cell with color
                    table.AddCell(new Cell()
                        .Add(new Paragraph(r.SentimentLabel)
                            .SetFont(boldFont)
                            .SetFontSize(8)
                            .SetFontColor(sentColor))
                        .SetBackgroundColor(rowBg)
                        .SetPadding(4));

                    table.AddCell(MakeCell(r.ConfidencePercent, normalFont, 8, rowBg));
                    table.AddCell(MakeCell(r.Status,            normalFont, 8, rowBg));

                    alt = !alt;
                }

                doc.Add(table);

                if (summary.Results.Count > 300)
                    doc.Add(new Paragraph(
                        $"Note: Showing first 300 of {summary.Results.Count} results.")
                        .SetFont(normalFont)
                        .SetFontSize(9)
                        .SetFontColor(grayColor));

                doc.Close();
            }
            catch (Exception ex)
            {
                throw new Exception($"PDF Export failed: {ex.Message}", ex);
            }
        }

        // Replace non-ASCII characters for PDF compatibility
        private static string CleanForPdf(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var sb = new System.Text.StringBuilder();
            foreach (char c in text)
                sb.Append(c >= 32 && c <= 126 ? c : '?');
            return sb.ToString();
        }

        // Helper: normal text cell
        private static Cell MakeCell(string text, PdfFont font,
            float size, DeviceRgb bg)
        {
            return new Cell()
                .Add(new Paragraph(text ?? "")
                    .SetFont(font)
                    .SetFontSize(size))
                .SetBackgroundColor(bg)
                .SetPadding(4);
        }

        // Helper: summary row
        private static void AddSummaryRow(Table table,
            string key, string value,
            PdfFont boldFont, PdfFont normalFont)
        {
            table.AddCell(new Cell()
                .Add(new Paragraph(key)
                    .SetFont(boldFont).SetFontSize(10))
                .SetPadding(5));
            table.AddCell(new Cell()
                .Add(new Paragraph(value ?? "")
                    .SetFont(normalFont).SetFontSize(10))
                .SetPadding(5));
        }
    }
}
