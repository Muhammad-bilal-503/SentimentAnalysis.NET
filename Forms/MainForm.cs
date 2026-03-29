using SentimentAnalyzerPro.Models;
using SentimentAnalyzerPro.Services;
using System.Drawing.Drawing2D;

namespace SentimentAnalyzerPro.Forms
{
    public partial class MainForm : Form
    {
        private readonly SentimentService _sentimentService;

        // Colors
        private Color _darkBg = Color.FromArgb(18, 18, 30);
        private Color _darkPanel = Color.FromArgb(28, 28, 45);
        private Color _darkCard = Color.FromArgb(38, 38, 58);
        private Color _accent = Color.FromArgb(99, 102, 241);
        private Color _accentGreen = Color.FromArgb(34, 197, 94);
        private Color _accentRed = Color.FromArgb(239, 68, 68);
        private Color _textLight = Color.FromArgb(248, 248, 255);
        private Color _textMuted = Color.FromArgb(148, 163, 184);

        private bool _isDarkMode = true;
        private BulkAnalysisSummary? _lastBulkSummary;

        // Controls
        private TabControl tabMain = null!;
        private Panel panelHeader = null!;
        private Label lblTitle = null!;
        private Label lblSubtitle = null!;
        private Button btnToggleTheme = null!;
        private Panel panelStatus = null!;
        private Label lblModelStatus = null!;
        private ProgressBar progressModel = null!;

        // Tab 1 - Single Analysis
        private Panel panelSingleBg = null!;
        private RichTextBox txtInput = null!;
        private Button btnAnalyze = null!;
        private Button btnClear = null!;
        private Panel panelResult = null!;
        private Label lblResultSentiment = null!;
        private Label lblResultConfidence = null!;
        private Label lblResultLanguage = null!;
        private Label lblResultTranslated = null!;
        private PictureBox picSentimentIcon = null!;

        // Tab 2 - Bulk Analysis
        private Panel panelBulkBg = null!;
        private Label lblSelectedFile = null!;
        private Button btnBrowseCSV = null!;
        private Button btnAnalyzeBulk = null!;
        private ProgressBar progressBulk = null!;
        private Label lblBulkProgress = null!;
        private Panel panelStats = null!;
        private Label lblTotal = null!;
        private Label lblPositive = null!;
        private Label lblNegative = null!;
        private Label lblSkipped = null!;
        private Label lblMood = null!;
        private Label lblTopPos = null!;
        private Label lblTopNeg = null!;
        private DataGridView gridResults = null!;

        // Tab 3 - Charts
        private Panel panelChartsBg = null!;
        private Panel panelPieChart = null!;
        private Panel panelBarChart = null!;

        // Tab 4 - Export
        private Panel panelExportBg = null!;
        private Button btnExportCSV = null!;
        private Button btnExportPDF = null!;
        private RichTextBox txtExportLog = null!;

        public MainForm()
        {
            _sentimentService = new SentimentService();
            InitializeComponent();
            ApplyTheme();
            InitModelAsync();
        }

        private void InitializeComponent()
        {
            this.Text = "Sentiment Analyzer Pro";
            this.Size = new Size(1100, 780);
            this.MinimumSize = new Size(900, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9.5f);
            this.BackColor = _darkBg;

            BuildHeader();
            BuildStatusBar();
            BuildTabControl();

            this.Controls.Add(tabMain);
            this.Controls.Add(panelStatus);
            this.Controls.Add(panelHeader);
        }

        private void BuildHeader()
        {
            panelHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                BackColor = _darkPanel,
                Padding = new Padding(20, 0, 20, 0)
            };

            lblTitle = new Label
            {
                Text = "🧠 Sentiment Analyzer Pro",
                Font = new Font("Segoe UI", 20f, FontStyle.Bold),
                ForeColor = _textLight,
                Location = new Point(20, 12),
                AutoSize = true
            };

            lblSubtitle = new Label
            {
                Text = "Multi-Language AI-Powered Sentiment Analysis  •  Powered by ML.NET",
                Font = new Font("Segoe UI", 9f),
                ForeColor = _textMuted,
                Location = new Point(22, 50),
                AutoSize = true
            };

            btnToggleTheme = new Button
            {
                Text = "☀ Light Mode",
                Size = new Size(120, 35),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Location = new Point(this.Width - 160, 22),
                FlatStyle = FlatStyle.Flat,
                ForeColor = _textLight,
                BackColor = _darkCard,
                Cursor = Cursors.Hand
            };
            btnToggleTheme.FlatAppearance.BorderColor = _accent;
            btnToggleTheme.Click += BtnToggleTheme_Click;

            panelHeader.Controls.AddRange(new Control[] { lblTitle, lblSubtitle, btnToggleTheme });
            panelHeader.Resize += (s, e) =>
            {
                btnToggleTheme.Location = new Point(panelHeader.Width - 150, 22);
            };
        }

        private void BuildStatusBar()
        {
            panelStatus = new Panel
            {
                Dock = DockStyle.Top,
                Height = 45,
                BackColor = _darkCard,
                Padding = new Padding(20, 5, 20, 5)
            };
            panelStatus.Location = new Point(0, 80);

            progressModel = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Location = new Point(20, 12),
                Size = new Size(200, 20),
                BackColor = _darkBg
            };

            lblModelStatus = new Label
            {
                Text = "⏳ Initializing model...",
                ForeColor = _accent,
                Font = new Font("Segoe UI", 9f),
                Location = new Point(230, 13),
                AutoSize = true
            };

            panelStatus.Controls.AddRange(new Control[] { progressModel, lblModelStatus });
        }

        private void BuildTabControl()
        {
            tabMain = new TabControl
            {
                Dock = DockStyle.Fill,
                Padding = new Point(15, 8),
                Font = new Font("Segoe UI", 10f)
            };
            tabMain.Location = new Point(0, 125);

            var tab1 = new TabPage("📝  Single Analysis");
            var tab2 = new TabPage("📂  Bulk CSV Analysis");
            var tab3 = new TabPage("📊  Charts");
            var tab4 = new TabPage("💾  Export");

            BuildTab1(tab1);
            BuildTab2(tab2);
            BuildTab3(tab3);
            BuildTab4(tab4);

            tabMain.TabPages.AddRange(new[] { tab1, tab2, tab3, tab4 });
        }

        private void BuildTab1(TabPage tab)
        {
            panelSingleBg = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            // Input group
            var lblInputTitle = MakeLabel("Enter Text to Analyze:", 0, 10, 16f, FontStyle.Bold);
            var lblInputHint = MakeLabel("Supports all languages — auto-detect & translate enabled 🌍", 0, 38, 9f);
            lblInputHint.ForeColor = _textMuted;

            txtInput = new RichTextBox
            {
                Location = new Point(0, 65),
                Size = new Size(900, 140),
                BackColor = _darkCard,
                ForeColor = _textLight,
                Font = new Font("Segoe UI", 11f),
                BorderStyle = BorderStyle.None,
                Padding = new Padding(10),
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            btnAnalyze = MakeButton("🔍  Analyze Sentiment", 0, 220, _accent);
            btnAnalyze.Size = new Size(200, 45);
            btnAnalyze.Click += BtnAnalyze_Click;

            btnClear = MakeButton("✕  Clear", 215, 220, _darkCard);
            btnClear.Size = new Size(100, 45);
            btnClear.Click += (s, e) => { txtInput.Clear(); ResetResultPanel(); };

            // Result panel
            panelResult = new Panel
            {
                Location = new Point(0, 285),
                Size = new Size(900, 180),
                BackColor = _darkCard,
                Visible = false
            };
            RoundPanel(panelResult);

            picSentimentIcon = new PictureBox
            {
                Location = new Point(20, 20),
                Size = new Size(80, 80),
                SizeMode = PictureBoxSizeMode.StretchImage
            };

            lblResultSentiment = new Label
            {
                Location = new Point(115, 20),
                Size = new Size(300, 45),
                Font = new Font("Segoe UI", 22f, FontStyle.Bold),
                ForeColor = _accentGreen
            };

            lblResultConfidence = new Label
            {
                Location = new Point(115, 70),
                Size = new Size(300, 28),
                Font = new Font("Segoe UI", 13f),
                ForeColor = _textMuted
            };

            lblResultLanguage = new Label
            {
                Location = new Point(115, 100),
                Size = new Size(400, 25),
                Font = new Font("Segoe UI", 9.5f),
                ForeColor = _textMuted
            };

            lblResultTranslated = new Label
            {
                Location = new Point(20, 140),
                Size = new Size(860, 30),
                Font = new Font("Segoe UI", 9f, FontStyle.Italic),
                ForeColor = _textMuted
            };

            panelResult.Controls.AddRange(new Control[] {
                picSentimentIcon, lblResultSentiment,
                lblResultConfidence, lblResultLanguage, lblResultTranslated
            });

            panelSingleBg.Controls.AddRange(new Control[] {
                lblInputTitle, lblInputHint, txtInput,
                btnAnalyze, btnClear, panelResult
            });

            tab.Controls.Add(panelSingleBg);
        }

        private void BuildTab2(TabPage tab)
        {
            panelBulkBg = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            // File selection
            var lblFileTitle = MakeLabel("Select CSV File with Comments:", 0, 10, 14f, FontStyle.Bold);

            lblSelectedFile = MakeLabel("No file selected...", 0, 42);
            lblSelectedFile.ForeColor = _textMuted;
            lblSelectedFile.Size = new Size(600, 25);

            btnBrowseCSV = MakeButton("📁  Browse CSV", 0, 70, _darkCard);
            btnBrowseCSV.Size = new Size(160, 40);
            btnBrowseCSV.Click += BtnBrowseCSV_Click;

            btnAnalyzeBulk = MakeButton("▶  Analyze All", 175, 70, _accent);
            btnAnalyzeBulk.Size = new Size(160, 40);
            btnAnalyzeBulk.Enabled = false;
            btnAnalyzeBulk.Click += BtnAnalyzeBulk_Click;

            progressBulk = new ProgressBar
            {
                Location = new Point(0, 125),
                Size = new Size(900, 22),
                BackColor = _darkCard,
                ForeColor = _accent,
                Visible = false
            };

            lblBulkProgress = MakeLabel("", 0, 150);
            lblBulkProgress.ForeColor = _textMuted;

            // Stats panel
            panelStats = new Panel
            {
                Location = new Point(0, 175),
                Size = new Size(900, 130),
                BackColor = _darkCard,
                Visible = false
            };
            RoundPanel(panelStats);

            lblTotal = MakeStat("Total", "0", 20, panelStats);
            lblPositive = MakeStat("✅ Positive", "0", 200, panelStats);
            lblNegative = MakeStat("❌ Negative", "0", 380, panelStats);
            lblSkipped = MakeStat("⚠ Skipped", "0", 560, panelStats);
            lblMood = MakeStat("Overall Mood", "—", 720, panelStats);

            lblTopPos = new Label
            {
                Location = new Point(20, 90),
                Size = new Size(400, 25),
                Font = new Font("Segoe UI", 9f),
                ForeColor = _textMuted
            };
            lblTopNeg = new Label
            {
                Location = new Point(450, 90),
                Size = new Size(400, 25),
                Font = new Font("Segoe UI", 9f),
                ForeColor = _textMuted
            };
            panelStats.Controls.AddRange(new Control[] { lblTopPos, lblTopNeg });

            // Results grid
            gridResults = new DataGridView
            {
                Location = new Point(0, 315),
                Size = new Size(900, 280),
                BackgroundColor = _darkCard,
                ForeColor = _textLight,
                GridColor = _darkPanel,
                BorderStyle = BorderStyle.None,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _darkPanel,
                    ForeColor = _textLight,
                    Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                    Padding = new Padding(5)
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _darkCard,
                    ForeColor = _textLight,
                    SelectionBackColor = _accent,
                    SelectionForeColor = Color.White
                },
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = true,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                Visible = false
            };

            gridResults.Columns.Add("text", "Comment");
            gridResults.Columns.Add("lang", "Language");
            gridResults.Columns.Add("sentiment", "Sentiment");
            gridResults.Columns.Add("confidence", "Confidence");
            gridResults.Columns.Add("status", "Status");
            gridResults.Columns[0].FillWeight = 45;
            gridResults.Columns[1].FillWeight = 15;
            gridResults.Columns[2].FillWeight = 15;
            gridResults.Columns[3].FillWeight = 15;
            gridResults.Columns[4].FillWeight = 10;

            panelBulkBg.Controls.AddRange(new Control[] {
                lblFileTitle, lblSelectedFile, btnBrowseCSV, btnAnalyzeBulk,
                progressBulk, lblBulkProgress, panelStats, gridResults
            });

            tab.Controls.Add(panelBulkBg);
        }

        private void BuildTab3(TabPage tab)
        {
            panelChartsBg = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            var lblChartTitle = MakeLabel("Analysis Charts", 0, 10, 16f, FontStyle.Bold);
            var lblChartHint = MakeLabel("Run a Bulk CSV Analysis first to see charts here.", 0, 42);
            lblChartHint.ForeColor = _textMuted;
            lblChartHint.Name = "lblChartHint";

            panelPieChart = new Panel
            {
                Location = new Point(0, 75),
                Size = new Size(430, 430),
                BackColor = _darkCard
            };
            RoundPanel(panelPieChart);
            panelPieChart.Paint += PanelPieChart_Paint;

            panelBarChart = new Panel
            {
                Location = new Point(450, 75),
                Size = new Size(430, 430),
                BackColor = _darkCard
            };
            RoundPanel(panelBarChart);
            panelBarChart.Paint += PanelBarChart_Paint;

            panelChartsBg.Controls.AddRange(new Control[] {
                lblChartTitle, lblChartHint, panelPieChart, panelBarChart
            });

            tab.Controls.Add(panelChartsBg);
        }

        private void BuildTab4(TabPage tab)
        {
            panelExportBg = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            var lblExportTitle = MakeLabel("Export Results", 0, 10, 16f, FontStyle.Bold);
            var lblHint = MakeLabel("Export your analysis results in different formats.", 0, 42);
            lblHint.ForeColor = _textMuted;

            btnExportCSV = MakeButton("📄  Export to CSV", 0, 80, _accentGreen);
            btnExportCSV.Size = new Size(200, 50);
            btnExportCSV.Click += BtnExportCSV_Click;

            btnExportPDF = MakeButton("📋  Export PDF Report", 220, 80, _accent);
            btnExportPDF.Size = new Size(200, 50);
            btnExportPDF.Click += BtnExportPDF_Click;

            txtExportLog = new RichTextBox
            {
                Location = new Point(0, 155),
                Size = new Size(900, 200),
                BackColor = _darkCard,
                ForeColor = _accentGreen,
                Font = new Font("Consolas", 9.5f),
                BorderStyle = BorderStyle.None,
                ReadOnly = true
            };

            panelExportBg.Controls.AddRange(new Control[] {
                lblExportTitle, lblHint, btnExportCSV, btnExportPDF, txtExportLog
            });

            tab.Controls.Add(panelExportBg);
        }

        // ===================== EVENT HANDLERS =====================

        private async void InitModelAsync()
        {
            var progress = new Progress<string>(msg =>
            {
                lblModelStatus.Text = $"⏳ {msg}";
            });

            var result = await _sentimentService.InitializeAsync(progress);

            progressModel.Visible = false;
            if (_sentimentService.IsReady)
            {
                lblModelStatus.Text = "✅ Model ready!  " + result;
                lblModelStatus.ForeColor = _accentGreen;
            }
            else
            {
                lblModelStatus.Text = "❌ " + result;
                lblModelStatus.ForeColor = _accentRed;
            }
        }

        private async void BtnAnalyze_Click(object? sender, EventArgs e)
        {
            if (!_sentimentService.IsReady)
            {
                MessageBox.Show("Model is still loading. Please wait.", "Not Ready", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string text = txtInput.Text.Trim();
            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show("Please enter some text to analyze.", "Empty Input", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            btnAnalyze.Enabled = false;
            btnAnalyze.Text = "⏳  Analyzing...";
            panelResult.Visible = false;

            var result = await _sentimentService.AnalyzeSingleAsync(text);

            ShowSingleResult(result);

            btnAnalyze.Enabled = true;
            btnAnalyze.Text = "🔍  Analyze Sentiment";
        }

        private void ShowSingleResult(AnalysisResult result)
        {
            if (result.Status == "Skipped")
            {
                lblResultSentiment.Text = "⚠ SKIPPED";
                lblResultSentiment.ForeColor = Color.Orange;
                lblResultConfidence.Text = result.StatusReason;
                lblResultLanguage.Text = "";
                lblResultTranslated.Text = "";
            }
            else
            {
                lblResultSentiment.Text = result.IsPositive ? "✅ POSITIVE" : "❌ NEGATIVE";
                lblResultSentiment.ForeColor = result.IsPositive ? _accentGreen : _accentRed;
                lblResultConfidence.Text = $"Confidence: {result.ConfidencePercent}";

                string langName = TranslationService.LanguageNames.TryGetValue(result.DetectedLanguage, out var n) ? n : result.DetectedLanguage;
                lblResultLanguage.Text = $"Detected Language: {langName}";

                if (result.DetectedLanguage != "en")
                    lblResultTranslated.Text = $"Translated: \"{result.TranslatedText}\"";
                else
                    lblResultTranslated.Text = "";
            }

            panelResult.Visible = true;
        }

        private void BtnBrowseCSV_Click(object? sender, EventArgs e)
        {
            using var dialog = new OpenFileDialog
            {
                Title = "Select CSV file with comments",
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                lblSelectedFile.Text = dialog.FileName;
                lblSelectedFile.ForeColor = _textLight;
                btnAnalyzeBulk.Enabled = true;
                btnAnalyzeBulk.Tag = dialog.FileName;
            }
        }

        private async void BtnAnalyzeBulk_Click(object? sender, EventArgs e)
        {
            if (!_sentimentService.IsReady)
            {
                MessageBox.Show("Model is still loading.", "Not Ready", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string? csvPath = btnAnalyzeBulk.Tag?.ToString();
            if (string.IsNullOrEmpty(csvPath) || !File.Exists(csvPath))
            {
                MessageBox.Show("Please select a valid CSV file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            btnAnalyzeBulk.Enabled = false;
            btnBrowseCSV.Enabled = false;
            progressBulk.Visible = true;
            progressBulk.Style = ProgressBarStyle.Continuous;
            progressBulk.Value = 0;
            gridResults.Rows.Clear();
            gridResults.Visible = false;
            panelStats.Visible = false;

            var progress = new Progress<(int current, int total, string message)>(p =>
            {
                lblBulkProgress.Text = p.message;
                if (p.total > 0)
                {
                    progressBulk.Maximum = p.total;
                    progressBulk.Value = Math.Min(p.current, p.total);
                }
            });

            _lastBulkSummary = await _sentimentService.AnalyzeBulkAsync(csvPath, progress);

            ShowBulkResults(_lastBulkSummary);

            btnAnalyzeBulk.Enabled = true;
            btnBrowseCSV.Enabled = true;
            progressBulk.Visible = false;
            lblBulkProgress.Text = "✅ Analysis complete!";
        }

        private void ShowBulkResults(BulkAnalysisSummary summary)
        {
            // Stats panel
            UpdateStat(lblTotal, summary.TotalComments.ToString());
            UpdateStat(lblPositive, $"{summary.PositiveCount} ({summary.PositivePercent:F0}%)");
            UpdateStat(lblNegative, $"{summary.NegativeCount} ({summary.NegativePercent:F0}%)");
            UpdateStat(lblSkipped, summary.SkippedCount.ToString());
            UpdateStat(lblMood, summary.OverallMood);
            lblTopPos.Text = $"🔝 Top Positive Word: \"{summary.TopPositiveWord}\"";
            lblTopNeg.Text = $"🔝 Top Negative Word: \"{summary.TopNegativeWord}\"";
            panelStats.Visible = true;

            // Grid
            gridResults.Rows.Clear();
            foreach (var r in summary.Results)
            {
                string langName = TranslationService.LanguageNames.TryGetValue(r.DetectedLanguage, out var n) ? n : r.DetectedLanguage;
                string shortText = r.OriginalText.Length > 60 ? r.OriginalText[..60] + "..." : r.OriginalText;
                gridResults.Rows.Add(shortText, langName, r.SentimentLabel, r.ConfidencePercent, r.Status);

                // Color rows
                var row = gridResults.Rows[gridResults.Rows.Count - 1];
                if (r.Status == "Analyzed")
                    row.DefaultCellStyle.ForeColor = r.IsPositive ? _accentGreen : _accentRed;
                else
                    row.DefaultCellStyle.ForeColor = Color.Orange;
            }
            gridResults.Visible = true;

            // Refresh charts
            panelPieChart.Invalidate();
            panelBarChart.Invalidate();
        }

        // ===================== CHARTS =====================

        private void PanelPieChart_Paint(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var panel = (Panel)sender!;

            if (_lastBulkSummary == null || _lastBulkSummary.AnalyzedCount == 0)
            {
                DrawNoData(g, panel.ClientRectangle);
                return;
            }

            // Title
            using var titleFont = new Font("Segoe UI", 12f, FontStyle.Bold);
            g.DrawString("Sentiment Distribution", titleFont, new SolidBrush(_textLight), new PointF(20, 15));

            float total = _lastBulkSummary.AnalyzedCount;
            float posAngle = (float)(_lastBulkSummary.PositiveCount / total * 360);

            var rect = new RectangleF(60, 55, 280, 280);

            // Pie slices
            using (var posBrush = new SolidBrush(_accentGreen))
                g.FillPie(posBrush, rect.X, rect.Y, rect.Width, rect.Height, -90, posAngle);

            using (var negBrush = new SolidBrush(_accentRed))
                g.FillPie(negBrush, rect.X, rect.Y, rect.Width, rect.Height, -90 + posAngle, 360 - posAngle);

            // Center hole (donut)
            using (var holeBrush = new SolidBrush(_darkCard))
                g.FillEllipse(holeBrush, rect.X + 70, rect.Y + 70, rect.Width - 140, rect.Height - 140);

            // Center text
            using var centerFont = new Font("Segoe UI", 14f, FontStyle.Bold);
            string mood = _lastBulkSummary.PositivePercent >= 50 ? "😊" : "😞";
            var centerStr = $"{mood}\n{_lastBulkSummary.PositivePercent:F0}% Pos";
            g.DrawString(centerStr, centerFont, new SolidBrush(_textLight),
                new RectangleF(rect.X + 75, rect.Y + 100, 140, 80),
                new StringFormat { Alignment = StringAlignment.Center });

            // Legend
            using var legendFont = new Font("Segoe UI", 10f);
            g.FillRectangle(new SolidBrush(_accentGreen), 20, 360, 14, 14);
            g.DrawString($"Positive: {_lastBulkSummary.PositiveCount}", legendFont, new SolidBrush(_textLight), 40, 357);

            g.FillRectangle(new SolidBrush(_accentRed), 200, 360, 14, 14);
            g.DrawString($"Negative: {_lastBulkSummary.NegativeCount}", legendFont, new SolidBrush(_textLight), 220, 357);
        }

        private void PanelBarChart_Paint(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var panel = (Panel)sender!;

            if (_lastBulkSummary == null || _lastBulkSummary.AnalyzedCount == 0)
            {
                DrawNoData(g, panel.ClientRectangle);
                return;
            }

            // Title
            using var titleFont = new Font("Segoe UI", 12f, FontStyle.Bold);
            g.DrawString("Comment Counts", titleFont, new SolidBrush(_textLight), new PointF(20, 15));

            float maxVal = Math.Max(_lastBulkSummary.PositiveCount, _lastBulkSummary.NegativeCount);
            if (maxVal == 0) return;

            float chartH = 280f;
            float baseY = 360f;
            float barW = 100f;

            // Positive bar
            float posH = (_lastBulkSummary.PositiveCount / maxVal) * chartH;
            using (var br = new LinearGradientBrush(
                new PointF(80, baseY - posH), new PointF(80, baseY),
                _accentGreen, Color.FromArgb(20, 200, 80)))
            {
                g.FillRectangle(br, 80, baseY - posH, barW, posH);
            }
            using var valFont = new Font("Segoe UI", 11f, FontStyle.Bold);
            g.DrawString(_lastBulkSummary.PositiveCount.ToString(), valFont,
                new SolidBrush(_accentGreen), new PointF(115, baseY - posH - 28));

            // Negative bar
            float negH = (_lastBulkSummary.NegativeCount / maxVal) * chartH;
            using (var br = new LinearGradientBrush(
                new PointF(240, baseY - negH), new PointF(240, baseY),
                _accentRed, Color.FromArgb(200, 30, 30)))
            {
                g.FillRectangle(br, 240, baseY - negH, barW, negH);
            }
            g.DrawString(_lastBulkSummary.NegativeCount.ToString(), valFont,
                new SolidBrush(_accentRed), new PointF(275, baseY - negH - 28));

            // Baseline
            using var basePen = new Pen(_textMuted, 1.5f);
            g.DrawLine(basePen, 60, baseY, 380, baseY);

            // Labels
            using var labelFont = new Font("Segoe UI", 10f);
            g.DrawString("Positive", labelFont, new SolidBrush(_accentGreen), 95, baseY + 10);
            g.DrawString("Negative", labelFont, new SolidBrush(_accentRed), 252, baseY + 10);

            // Skipped bar (if any)
            if (_lastBulkSummary.SkippedCount > 0)
            {
                float skipH = (_lastBulkSummary.SkippedCount / maxVal) * chartH;
                g.FillRectangle(new SolidBrush(Color.Orange), 400, baseY - skipH, barW, skipH);
                g.DrawString(_lastBulkSummary.SkippedCount.ToString(), valFont,
                    new SolidBrush(Color.Orange), new PointF(435, baseY - skipH - 28));
                g.DrawString("Skipped", labelFont, new SolidBrush(Color.Orange), 412, baseY + 10);
            }
        }

        private void DrawNoData(Graphics g, Rectangle rect)
        {
            using var font = new Font("Segoe UI", 12f);
            g.DrawString("No data yet.\nRun Bulk Analysis first.",
                font, new SolidBrush(_textMuted),
                new RectangleF(0, 0, rect.Width, rect.Height),
                new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
        }

        // ===================== EXPORT =====================

        private async void BtnExportCSV_Click(object? sender, EventArgs e)
        {
            if (_lastBulkSummary == null)
            {
                MessageBox.Show("Please run a Bulk Analysis first.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var dialog = new SaveFileDialog
            {
                Title = "Save CSV Results",
                Filter = "CSV Files (*.csv)|*.csv",
                FileName = $"sentiment_results_{DateTime.Now:yyyyMMdd_HHmm}.csv"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                await _sentimentService.ExportToCsvAsync(_lastBulkSummary, dialog.FileName);
                txtExportLog.Text += $"[{DateTime.Now:HH:mm:ss}] ✅ CSV exported: {dialog.FileName}\n";
                MessageBox.Show("CSV exported successfully!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnExportPDF_Click(object? sender, EventArgs e)
        {
            if (_lastBulkSummary == null)
            {
                MessageBox.Show("Please run a Bulk Analysis first.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var dialog = new SaveFileDialog
            {
                Title = "Save PDF Report",
                Filter = "PDF Files (*.pdf)|*.pdf",
                FileName = $"sentiment_report_{DateTime.Now:yyyyMMdd_HHmm}.pdf"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                PdfExporter.Export(_lastBulkSummary, dialog.FileName);
                txtExportLog.Text += $"[{DateTime.Now:HH:mm:ss}] ✅ PDF exported: {dialog.FileName}\n";
                MessageBox.Show("PDF report exported successfully!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // ===================== THEME =====================

        private void BtnToggleTheme_Click(object? sender, EventArgs e)
        {
            _isDarkMode = !_isDarkMode;
            ApplyTheme();
            btnToggleTheme.Text = _isDarkMode ? "☀ Light Mode" : "🌙 Dark Mode";
        }

        private void ApplyTheme()
        {
            if (_isDarkMode)
            {
                _darkBg = Color.FromArgb(18, 18, 30);
                _darkPanel = Color.FromArgb(28, 28, 45);
                _darkCard = Color.FromArgb(38, 38, 58);
                _textLight = Color.FromArgb(248, 248, 255);
                _textMuted = Color.FromArgb(148, 163, 184);
            }
            else
            {
                _darkBg = Color.FromArgb(245, 247, 250);
                _darkPanel = Color.FromArgb(255, 255, 255);
                _darkCard = Color.FromArgb(230, 235, 245);
                _textLight = Color.FromArgb(20, 20, 40);
                _textMuted = Color.FromArgb(100, 110, 130);
            }

            this.BackColor = _darkBg;
            if (panelHeader != null) panelHeader.BackColor = _darkPanel;
            if (panelStatus != null) panelStatus.BackColor = _darkCard;
            if (txtInput != null) { txtInput.BackColor = _darkCard; txtInput.ForeColor = _textLight; }
            if (gridResults != null)
            {
                gridResults.BackgroundColor = _darkCard;
                gridResults.DefaultCellStyle.BackColor = _darkCard;
                gridResults.DefaultCellStyle.ForeColor = _textLight;
                gridResults.ColumnHeadersDefaultCellStyle.BackColor = _darkPanel;
                gridResults.ColumnHeadersDefaultCellStyle.ForeColor = _textLight;
            }
            if (panelResult != null) panelResult.BackColor = _darkCard;
            if (panelStats != null) panelStats.BackColor = _darkCard;
            if (panelPieChart != null) { panelPieChart.BackColor = _darkCard; panelPieChart.Invalidate(); }
            if (panelBarChart != null) { panelBarChart.BackColor = _darkCard; panelBarChart.Invalidate(); }

            this.Invalidate(true);
        }

        // ===================== HELPERS =====================

        private void ResetResultPanel()
        {
            panelResult.Visible = false;
        }

        private Label MakeLabel(string text, int x, int y, float size = 9.5f, FontStyle style = FontStyle.Regular)
        {
            return new Label
            {
                Text = text,
                Location = new Point(x, y),
                Font = new Font("Segoe UI", size, style),
                ForeColor = _textLight,
                AutoSize = true
            };
        }

        private Button MakeButton(string text, int x, int y, Color backColor)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                FlatStyle = FlatStyle.Flat,
                BackColor = backColor,
                ForeColor = _textLight,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = ControlPaint.Light(backColor, 0.2f);
            return btn;
        }

        private Label MakeStat(string label, string value, int x, Panel parent)
        {
            var lblLabel = new Label
            {
                Text = label,
                Location = new Point(x, 12),
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = _textMuted,
                AutoSize = true
            };
            var lblValue = new Label
            {
                Text = value,
                Location = new Point(x, 33),
                Font = new Font("Segoe UI", 18f, FontStyle.Bold),
                ForeColor = _textLight,
                AutoSize = true,
                Name = $"val_{x}"
            };
            parent.Controls.AddRange(new Control[] { lblLabel, lblValue });
            return lblValue;
        }

        private void UpdateStat(Label lbl, string value)
        {
            lbl.Text = value;
        }

        private void RoundPanel(Panel panel)
        {
            panel.Paint += (s, e) =>
            {
                var p = (Panel)s!;
                using var pen = new Pen(_accent, 1.5f);
                var rect = new Rectangle(0, 0, p.Width - 1, p.Height - 1);
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.DrawRoundedRectangle(pen, rect, 10);
            };
        }
    }

    // Extension for rounded rectangles
    public static class GraphicsExtensions
    {
        public static void DrawRoundedRectangle(this Graphics g, Pen pen, Rectangle rect, int radius)
        {
            using var path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddArc(rect.X, rect.Y, radius * 2, radius * 2, 180, 90);
            path.AddArc(rect.Right - radius * 2, rect.Y, radius * 2, radius * 2, 270, 90);
            path.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(rect.X, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseFigure();
            g.DrawPath(pen, path);
        }
    }
}
