using SentimentAnalyzerPro.Models;
using SentimentAnalyzerPro.Services;
using System.Drawing.Drawing2D;

namespace SentimentAnalyzerPro.Forms
{
    public partial class MainForm : Form
    {
        private readonly SentimentService _sentimentService;
        private bool _isDarkMode = true;
        private BulkAnalysisSummary? _lastBulkSummary;

        // Theme colors
        private Color BgColor      => _isDarkMode ? Color.FromArgb(15, 15, 25)    : Color.FromArgb(245, 247, 252);
        private Color PanelColor   => _isDarkMode ? Color.FromArgb(25, 25, 40)    : Color.FromArgb(255, 255, 255);
        private Color CardColor    => _isDarkMode ? Color.FromArgb(35, 35, 55)    : Color.FromArgb(235, 238, 248);
        private Color AccentColor  => Color.FromArgb(99, 102, 241);
        private Color GreenColor   => Color.FromArgb(34, 197, 94);
        private Color RedColor     => Color.FromArgb(239, 68, 68);
        private Color TextColor    => _isDarkMode ? Color.FromArgb(240, 240, 255) : Color.FromArgb(20, 20, 40);
        private Color MutedColor   => _isDarkMode ? Color.FromArgb(140, 150, 180) : Color.FromArgb(100, 110, 140);
        private Color BorderColor  => _isDarkMode ? Color.FromArgb(60, 60, 90)    : Color.FromArgb(200, 205, 220);

        // Controls
        private Panel pnlHeader = null!;
        private Label lblTitle = null!, lblSub = null!;
        private Button btnTheme = null!;
        private Panel pnlStatus = null!;
        private Label lblStatus = null!;
        private ProgressBar pbModel = null!;
        private TabControl tcMain = null!;

        // Tab1
        private Panel pnl1 = null!;
        private RichTextBox rtbInput = null!;
        private Button btnAnalyze = null!, btnClear = null!;
        private Panel pnlResult = null!;
        private Label lblSentiment = null!, lblConf = null!, lblLang = null!, lblTrans = null!;

        // Tab2
        private Panel pnl2 = null!;
        private Label lblFile = null!;
        private Button btnBrowse = null!, btnAnalyzeBulk = null!;
        private ProgressBar pbBulk = null!;
        private Label lblProgress = null!;
        private Panel pnlStats = null!;
        private Label lblTotal = null!, lblPos = null!, lblNeg = null!, lblSkip = null!, lblMood = null!;
        private Label lblTopPos = null!, lblTopNeg = null!;
        private DataGridView dgvResults = null!;

        // Tab3
        private Panel pnlPie = null!, pnlBar = null!;

        // Tab4
        private Button btnExportCsv = null!, btnExportPdf = null!;
        private RichTextBox rtbLog = null!;

        public MainForm()
        {
            _sentimentService = new SentimentService();
            InitializeComponent();
            InitModelAsync();
        }

        private void InitializeComponent()
        {
            Text = "Sentiment Analyzer Pro";
            Size = new Size(1150, 820);
            MinimumSize = new Size(950, 700);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9.5f);
            BackColor = BgColor;
            DoubleBuffered = true;

            // IMPORTANT: Build in this order for correct Dock layout
            // Fill must be added FIRST, then Top panels on top of it
            BuildTabs();       // Fill - add first
            BuildStatusBar();  // Top - add second
            BuildHeader();     // Top - add last (appears at very top)
        }

        // ============================================================
        // HEADER
        // ============================================================
        private void BuildHeader()
        {
            pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 75,
                BackColor = PanelColor,
                Padding = new Padding(20, 0, 20, 0)
            };

            lblTitle = new Label
            {
                Text = "🧠  Sentiment Analyzer Pro",
                Font = new Font("Segoe UI", 19f, FontStyle.Bold),
                ForeColor = TextColor,
                Location = new Point(20, 10),
                AutoSize = true
            };

            lblSub = new Label
            {
                Text = "Multi-Language AI Sentiment Analysis  •  Powered by ML.NET",
                Font = new Font("Segoe UI", 9f),
                ForeColor = MutedColor,
                Location = new Point(22, 47),
                AutoSize = true
            };

            btnTheme = new Button
            {
                Text = "☀  Light Mode",
                Size = new Size(130, 36),
                FlatStyle = FlatStyle.Flat,
                ForeColor = TextColor,
                BackColor = CardColor,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnTheme.FlatAppearance.BorderColor = BorderColor;
            btnTheme.FlatAppearance.BorderSize = 1;
            btnTheme.Click += (s, e) => ToggleTheme();

            pnlHeader.Controls.AddRange(new Control[] { lblTitle, lblSub, btnTheme });
            pnlHeader.Resize += (s, e) =>
                btnTheme.Location = new Point(pnlHeader.Width - 150, 19);
            btnTheme.Location = new Point(970, 19);

            Controls.Add(pnlHeader);
        }

        // ============================================================
        // STATUS BAR
        // ============================================================
        private void BuildStatusBar()
        {
            pnlStatus = new Panel
            {
                Height = 42,
                BackColor = CardColor,
                Dock = DockStyle.Top,
                Padding = new Padding(15, 0, 15, 0)
            };

            pbModel = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 25,
                Location = new Point(15, 12),
                Size = new Size(180, 18),
                BackColor = BgColor
            };

            lblStatus = new Label
            {
                Text = "⏳  Initializing model...",
                ForeColor = AccentColor,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Location = new Point(205, 13),
                AutoSize = true
            };

            pnlStatus.Controls.AddRange(new Control[] { pbModel, lblStatus });
            Controls.Add(pnlStatus);
        }

        // ============================================================
        // TABS
        // ============================================================
        private void BuildTabs()
        {
            tcMain = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10f),
                Padding = new Point(16, 8)
            };

            var t1 = new TabPage("  📝  Single Analysis  ");
            var t2 = new TabPage("  📂  Bulk CSV Analysis  ");
            var t3 = new TabPage("  📊  Charts  ");
            var t4 = new TabPage("  💾  Export  ");

            BuildTab1(t1);
            BuildTab2(t2);
            BuildTab3(t3);
            BuildTab4(t4);

            tcMain.TabPages.AddRange(new[] { t1, t2, t3, t4 });
            Controls.Add(tcMain);
        }

        // ============================================================
        // TAB 1 - SINGLE ANALYSIS
        // ============================================================
        private void BuildTab1(TabPage tab)
        {
            tab.BackColor = BgColor;
            pnl1 = new Panel { Dock = DockStyle.Fill, Padding = new Padding(25) };
            pnl1.BackColor = BgColor;

            var lbl1 = MkLabel("Enter Text to Analyze:", 0, 15, 15f, FontStyle.Bold);
            var lbl2 = MkLabel("🌍  Supports 30+ languages — auto-detect enabled", 2, 45, 9f);
            lbl2.ForeColor = MutedColor;

            rtbInput = new RichTextBox
            {
                Location = new Point(0, 72),
                Size = new Size(1060, 150),
                BackColor = CardColor,
                ForeColor = TextColor,
                Font = new Font("Segoe UI", 11f),
                BorderStyle = BorderStyle.None,
                Padding = new Padding(12),
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            btnAnalyze = MkBtn("🔍  Analyze Sentiment", 0, 238, AccentColor);
            btnAnalyze.Size = new Size(210, 46);
            btnAnalyze.Click += BtnAnalyze_Click;

            btnClear = MkBtn("✕  Clear", 220, 238, CardColor);
            btnClear.Size = new Size(100, 46);
            btnClear.FlatAppearance.BorderColor = BorderColor;
            btnClear.FlatAppearance.BorderSize = 1;
            btnClear.Click += (s, e) => { rtbInput.Clear(); pnlResult.Visible = false; };

            // Result panel
            pnlResult = new Panel
            {
                Location = new Point(0, 305),
                Size = new Size(1060, 180),
                BackColor = CardColor,
                Visible = false
            };

            lblSentiment = new Label
            {
                Location = new Point(20, 25),
                Size = new Size(500, 52),
                Font = new Font("Segoe UI", 26f, FontStyle.Bold),
                ForeColor = GreenColor
            };

            lblConf = new Label
            {
                Location = new Point(22, 82),
                Size = new Size(400, 30),
                Font = new Font("Segoe UI", 13f),
                ForeColor = MutedColor
            };

            lblLang = new Label
            {
                Location = new Point(22, 115),
                Size = new Size(600, 26),
                Font = new Font("Segoe UI", 10f),
                ForeColor = MutedColor
            };

            lblTrans = new Label
            {
                Location = new Point(22, 148),
                Size = new Size(1000, 24),
                Font = new Font("Segoe UI", 9f, FontStyle.Italic),
                ForeColor = MutedColor
            };

            pnlResult.Controls.AddRange(new Control[]
                { lblSentiment, lblConf, lblLang, lblTrans });

            pnl1.Controls.AddRange(new Control[]
                { lbl1, lbl2, rtbInput, btnAnalyze, btnClear, pnlResult });

            tab.Controls.Add(pnl1);
        }

        // ============================================================
        // TAB 2 - BULK ANALYSIS
        // ============================================================
        private void BuildTab2(TabPage tab)
        {
            tab.BackColor = BgColor;
            pnl2 = new Panel { Dock = DockStyle.Fill, Padding = new Padding(25) };
            pnl2.BackColor = BgColor;

            var lbl = MkLabel("Select CSV File with Comments:", 0, 12, 14f, FontStyle.Bold);

            lblFile = MkLabel("No file selected...", 0, 44, 9.5f);
            lblFile.ForeColor = MutedColor;
            lblFile.Size = new Size(700, 24);

            btnBrowse = MkBtn("📁  Browse CSV", 0, 72, CardColor);
            btnBrowse.Size = new Size(160, 42);
            btnBrowse.FlatAppearance.BorderColor = AccentColor;
            btnBrowse.FlatAppearance.BorderSize = 1;
            btnBrowse.Click += BtnBrowse_Click;

            btnAnalyzeBulk = MkBtn("▶  Analyze All", 172, 72, AccentColor);
            btnAnalyzeBulk.Size = new Size(160, 42);
            btnAnalyzeBulk.Enabled = false;
            btnAnalyzeBulk.Click += BtnAnalyzeBulk_Click;

            pbBulk = new ProgressBar
            {
                Location = new Point(0, 130),
                Size = new Size(1060, 20),
                ForeColor = AccentColor,
                BackColor = CardColor,
                Visible = false
            };

            lblProgress = MkLabel("", 0, 156, 9f);
            lblProgress.ForeColor = MutedColor;
            lblProgress.Size = new Size(700, 22);

            // Stats panel
            pnlStats = new Panel
            {
                Location = new Point(0, 180),
                Size = new Size(1060, 110),
                BackColor = CardColor,
                Visible = false
            };

            // Stat boxes
            lblTotal = MkStatBox("Total", "0", 15, pnlStats);
            lblPos   = MkStatBox("✅  Positive", "0", 225, pnlStats);
            lblNeg   = MkStatBox("❌  Negative", "0", 435, pnlStats);
            lblSkip  = MkStatBox("⚠  Skipped", "0", 645, pnlStats);
            lblMood  = MkStatBox("Overall Mood", "—", 855, pnlStats);

            lblTopPos = new Label
            {
                Location = new Point(15, 82),
                Size = new Size(500, 22),
                Font = new Font("Segoe UI", 9f),
                ForeColor = MutedColor
            };
            lblTopNeg = new Label
            {
                Location = new Point(535, 82),
                Size = new Size(500, 22),
                Font = new Font("Segoe UI", 9f),
                ForeColor = MutedColor
            };
            pnlStats.Controls.AddRange(new Control[] { lblTopPos, lblTopNeg });

            // Results grid
            dgvResults = new DataGridView
            {
                Location = new Point(0, 300),
                BackgroundColor = CardColor,
                ForeColor = TextColor,
                GridColor = BorderColor,
                BorderStyle = BorderStyle.None,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(50, 50, 80),
                    ForeColor = Color.White,
                    Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                    Padding = new Padding(6)
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = CardColor,
                    ForeColor = TextColor,
                    SelectionBackColor = AccentColor,
                    SelectionForeColor = Color.White,
                    Padding = new Padding(4)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _isDarkMode
                        ? Color.FromArgb(30, 30, 50)
                        : Color.FromArgb(242, 244, 255),
                    ForeColor = TextColor
                },
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = true,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                AllowUserToResizeRows = false,
                ColumnHeadersHeightSizeMode =
                    DataGridViewColumnHeadersHeightSizeMode.EnableResizing,
                ColumnHeadersHeight = 38,
                RowTemplate = { Height = 32 },
                Visible = false,
                ScrollBars = ScrollBars.Both
            };

            dgvResults.Columns.Add(new DataGridViewTextBoxColumn
                { Name = "col_text", HeaderText = "Comment", FillWeight = 42 });
            dgvResults.Columns.Add(new DataGridViewTextBoxColumn
                { Name = "col_lang", HeaderText = "Language", FillWeight = 13 });
            dgvResults.Columns.Add(new DataGridViewTextBoxColumn
                { Name = "col_sent", HeaderText = "Sentiment", FillWeight = 13 });
            dgvResults.Columns.Add(new DataGridViewTextBoxColumn
                { Name = "col_conf", HeaderText = "Confidence", FillWeight = 13 });
            dgvResults.Columns.Add(new DataGridViewTextBoxColumn
                { Name = "col_stat", HeaderText = "Status", FillWeight = 10 });

            // Auto resize grid with form
            pnl2.Resize += (s, e) =>
            {
                int w = pnl2.Width - 50;
                int h = pnl2.Height - 310;
                if (w > 200 && h > 100)
                    dgvResults.Size = new Size(w, h);
            };

            pnl2.Controls.AddRange(new Control[]
            {
                lbl, lblFile, btnBrowse, btnAnalyzeBulk,
                pbBulk, lblProgress, pnlStats, dgvResults
            });

            tab.Controls.Add(pnl2);
        }

        // ============================================================
        // TAB 3 - CHARTS
        // ============================================================
        private void BuildTab3(TabPage tab)
        {
            tab.BackColor = BgColor;
            var pnl = new Panel { Dock = DockStyle.Fill, Padding = new Padding(25) };
            pnl.BackColor = BgColor;

            var lbl = MkLabel("Sentiment Charts", 0, 12, 15f, FontStyle.Bold);
            var hint = MkLabel("Run Bulk CSV Analysis first to see charts.", 2, 44, 9f);
            hint.ForeColor = MutedColor;
            hint.Name = "lblChartHint";

            pnlPie = new Panel
            {
                Location = new Point(0, 75),
                Size = new Size(470, 470),
                BackColor = CardColor
            };
            pnlPie.Paint += PnlPie_Paint;

            pnlBar = new Panel
            {
                Location = new Point(490, 75),
                Size = new Size(470, 470),
                BackColor = CardColor
            };
            pnlBar.Paint += PnlBar_Paint;

            pnl.Controls.AddRange(new Control[] { lbl, hint, pnlPie, pnlBar });
            tab.Controls.Add(pnl);
        }

        // ============================================================
        // TAB 4 - EXPORT
        // ============================================================
        private void BuildTab4(TabPage tab)
        {
            tab.BackColor = BgColor;
            var pnl = new Panel { Dock = DockStyle.Fill, Padding = new Padding(25) };
            pnl.BackColor = BgColor;

            var lbl = MkLabel("Export Results", 0, 12, 15f, FontStyle.Bold);
            var hint = MkLabel("Export your analysis results after running Bulk CSV Analysis.", 2, 44, 9f);
            hint.ForeColor = MutedColor;

            // CSV button
            btnExportCsv = MkBtn("📄  Export to CSV", 0, 85, GreenColor);
            btnExportCsv.Size = new Size(210, 52);
            btnExportCsv.Font = new Font("Segoe UI", 10.5f, FontStyle.Bold);
            btnExportCsv.Click += BtnExportCsv_Click;

            // PDF button
            btnExportPdf = MkBtn("📋  Export PDF Report", 225, 85, AccentColor);
            btnExportPdf.Size = new Size(210, 52);
            btnExportPdf.Font = new Font("Segoe UI", 10.5f, FontStyle.Bold);
            btnExportPdf.Click += BtnExportPdf_Click;

            // Info box
            var infoPanel = new Panel
            {
                Location = new Point(0, 155),
                Size = new Size(700, 80),
                BackColor = CardColor
            };

            var infoLbl = new Label
            {
                Text = "ℹ  CSV Export: Saves all comments with language, sentiment and confidence.\r\n" +
                       "ℹ  PDF Report: Generates a professional formatted report with summary table.",
                Location = new Point(15, 12),
                Size = new Size(670, 55),
                Font = new Font("Segoe UI", 9.5f),
                ForeColor = MutedColor
            };
            infoPanel.Controls.Add(infoLbl);

            // Log
            rtbLog = new RichTextBox
            {
                Location = new Point(0, 255),
                Size = new Size(1060, 200),
                BackColor = CardColor,
                ForeColor = GreenColor,
                Font = new Font("Consolas", 9.5f),
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            pnl.Controls.AddRange(new Control[]
                { lbl, hint, btnExportCsv, btnExportPdf, infoPanel, rtbLog });
            tab.Controls.Add(pnl);
        }

        // ============================================================
        // MODEL INIT
        // ============================================================
        private async void InitModelAsync()
        {
            var progress = new Progress<string>(msg => lblStatus.Text = $"⏳  {msg}");
            var result = await _sentimentService.InitializeAsync(progress);

            pbModel.Visible = false;
            if (_sentimentService.IsReady)
            {
                lblStatus.Text = "✅  Model ready!  " + result;
                lblStatus.ForeColor = GreenColor;
            }
            else
            {
                lblStatus.Text = "❌  " + result;
                lblStatus.ForeColor = RedColor;
            }
        }

        // ============================================================
        // ANALYZE SINGLE
        // ============================================================
        private async void BtnAnalyze_Click(object? sender, EventArgs e)
        {
            if (!_sentimentService.IsReady)
            {
                MessageBox.Show("Model is still loading. Please wait.",
                    "Not Ready", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string text = rtbInput.Text.Trim();
            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show("Please enter some text.",
                    "Empty", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            btnAnalyze.Enabled = false;
            btnAnalyze.Text = "⏳  Analyzing...";
            pnlResult.Visible = false;

            var result = await _sentimentService.AnalyzeSingleAsync(text);

            // Show result
            if (result.Status == "Skipped")
            {
                lblSentiment.Text = "⚠  SKIPPED";
                lblSentiment.ForeColor = Color.Orange;
                lblConf.Text = result.StatusReason;
                lblLang.Text = "";
                lblTrans.Text = "";
            }
            else
            {
                lblSentiment.Text = result.IsPositive ? "✅  POSITIVE" : "❌  NEGATIVE";
                lblSentiment.ForeColor = result.IsPositive ? GreenColor : RedColor;
                lblConf.Text = $"Confidence:  {result.ConfidencePercent}";

                string langName = TranslationService.LanguageNames
                    .TryGetValue(result.DetectedLanguage, out var n) ? n : result.DetectedLanguage;
                lblLang.Text = $"Detected Language:  {langName}";

                lblTrans.Text = result.DetectedLanguage != "en" && !string.IsNullOrEmpty(result.TranslatedText)
                    ? $"Translated:  \"{result.TranslatedText}\""
                    : "";
            }

            pnlResult.Visible = true;
            btnAnalyze.Enabled = true;
            btnAnalyze.Text = "🔍  Analyze Sentiment";
        }

        // ============================================================
        // BULK BROWSE
        // ============================================================
        private void BtnBrowse_Click(object? sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog
            {
                Title = "Select CSV file",
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                lblFile.Text = dlg.FileName;
                lblFile.ForeColor = TextColor;
                btnAnalyzeBulk.Enabled = true;
                btnAnalyzeBulk.Tag = dlg.FileName;
            }
        }

        // ============================================================
        // BULK ANALYZE
        // ============================================================
        private async void BtnAnalyzeBulk_Click(object? sender, EventArgs e)
        {
            if (!_sentimentService.IsReady)
            {
                MessageBox.Show("Model is still loading.", "Not Ready",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string? path = btnAnalyzeBulk.Tag?.ToString();
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                MessageBox.Show("Please select a valid CSV file.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            btnAnalyzeBulk.Enabled = false;
            btnBrowse.Enabled = false;
            pbBulk.Visible = true;
            pbBulk.Value = 0;
            dgvResults.Rows.Clear();
            dgvResults.Visible = false;
            pnlStats.Visible = false;

            var progress = new Progress<(int current, int total, string message)>(p =>
            {
                lblProgress.Text = p.message;
                if (p.total > 0)
                {
                    pbBulk.Maximum = p.total;
                    pbBulk.Value = Math.Min(p.current, p.total);
                }
            });

            _lastBulkSummary = await _sentimentService.AnalyzeBulkAsync(path, progress);
            ShowBulkResults(_lastBulkSummary);

            btnAnalyzeBulk.Enabled = true;
            btnBrowse.Enabled = true;
            pbBulk.Visible = false;
            lblProgress.Text = $"✅  Analysis complete! {_lastBulkSummary.AnalyzedCount} comments analyzed.";
        }

        private void ShowBulkResults(BulkAnalysisSummary s)
        {
            lblTotal.Text = s.TotalComments.ToString();
            lblPos.Text = $"{s.PositiveCount}\n({s.PositivePercent:F0}%)";
            lblNeg.Text = $"{s.NegativeCount}\n({s.NegativePercent:F0}%)";
            lblSkip.Text = s.SkippedCount.ToString();
            lblMood.Text = s.PositivePercent >= 50 ? "😊 POSITIVE" : "😞 NEGATIVE";
            lblTopPos.Text = $"🔝 Top Positive Word: \"{s.TopPositiveWord}\"";
            lblTopNeg.Text = $"🔝 Top Negative Word: \"{s.TopNegativeWord}\"";
            pnlStats.Visible = true;

            dgvResults.Rows.Clear();
            dgvResults.SuspendLayout();

            foreach (var r in s.Results)
            {
                string lang = TranslationService.LanguageNames
                    .TryGetValue(r.DetectedLanguage, out var n) ? n : r.DetectedLanguage;
                string shortText = r.OriginalText.Length > 70
                    ? r.OriginalText[..70] + "..." : r.OriginalText;

                int idx = dgvResults.Rows.Add(shortText, lang, r.SentimentLabel,
                    r.ConfidencePercent, r.Status);

                var row = dgvResults.Rows[idx];
                if (r.Status == "Analyzed")
                    row.DefaultCellStyle.ForeColor = r.IsPositive ? GreenColor : RedColor;
                else
                    row.DefaultCellStyle.ForeColor = Color.Orange;
            }

            dgvResults.ResumeLayout();
            dgvResults.Visible = true;

            // Refresh charts
            pnlPie?.Invalidate();
            pnlBar?.Invalidate();
        }

        // ============================================================
        // CHARTS
        // ============================================================
        private void PnlPie_Paint(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var p = (Panel)sender!;

            if (_lastBulkSummary == null || _lastBulkSummary.AnalyzedCount == 0)
            {
                DrawNoData(g, p.ClientRectangle); return;
            }

            using var titleFont = new Font("Segoe UI", 13f, FontStyle.Bold);
            g.DrawString("Sentiment Distribution", titleFont,
                new SolidBrush(TextColor), new PointF(20, 18));

            float total = _lastBulkSummary.AnalyzedCount;
            float posAngle = (float)(_lastBulkSummary.PositiveCount / total * 360f);

            var rect = new RectangleF(55, 55, 300, 300);

            using (var b = new SolidBrush(GreenColor))
                g.FillPie(b, rect.X, rect.Y, rect.Width, rect.Height, -90, posAngle);
            using (var b = new SolidBrush(RedColor))
                g.FillPie(b, rect.X, rect.Y, rect.Width, rect.Height, -90 + posAngle, 360 - posAngle);
            using (var b = new SolidBrush(CardColor))
                g.FillEllipse(b, rect.X + 75, rect.Y + 75, rect.Width - 150, rect.Height - 150);

            // Center text
            using var cf = new Font("Segoe UI", 15f, FontStyle.Bold);
            string mood = _lastBulkSummary.PositivePercent >= 50 ? "😊" : "😞";
            g.DrawString($"{mood}", cf, new SolidBrush(TextColor),
                new RectangleF(rect.X + 95, rect.Y + 95, 120, 50),
                new StringFormat { Alignment = StringAlignment.Center });
            g.DrawString($"{_lastBulkSummary.PositivePercent:F1}% Pos", cf,
                new SolidBrush(GreenColor),
                new RectangleF(rect.X + 70, rect.Y + 145, 170, 40),
                new StringFormat { Alignment = StringAlignment.Center });

            // Legend
            using var lf = new Font("Segoe UI", 10f);
            g.FillRectangle(new SolidBrush(GreenColor), 20, 380, 14, 14);
            g.DrawString($"Positive: {_lastBulkSummary.PositiveCount}",
                lf, new SolidBrush(TextColor), 40, 377);
            g.FillRectangle(new SolidBrush(RedColor), 215, 380, 14, 14);
            g.DrawString($"Negative: {_lastBulkSummary.NegativeCount}",
                lf, new SolidBrush(TextColor), 235, 377);
        }

        private void PnlBar_Paint(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var p = (Panel)sender!;

            if (_lastBulkSummary == null || _lastBulkSummary.AnalyzedCount == 0)
            {
                DrawNoData(g, p.ClientRectangle); return;
            }

            using var titleFont = new Font("Segoe UI", 13f, FontStyle.Bold);
            g.DrawString("Comment Counts", titleFont,
                new SolidBrush(TextColor), new PointF(20, 18));

            float maxVal = Math.Max(
                Math.Max(_lastBulkSummary.PositiveCount, _lastBulkSummary.NegativeCount),
                Math.Max(_lastBulkSummary.SkippedCount, 1));

            float chartH = 300f, baseY = 390f, barW = 90f;

            void DrawBar(float x, float count, Color top, Color bottom, string label)
            {
                float h = (count / maxVal) * chartH;
                using var br = new LinearGradientBrush(
                    new PointF(x, baseY - h), new PointF(x, baseY), top, bottom);
                g.FillRectangle(br, x, baseY - h, barW, h);

                using var vf = new Font("Segoe UI", 11f, FontStyle.Bold);
                g.DrawString(((int)count).ToString(), vf, new SolidBrush(top),
                    new PointF(x + barW / 2 - 15, baseY - h - 28));

                using var lf = new Font("Segoe UI", 10f);
                g.DrawString(label, lf, new SolidBrush(TextColor),
                    new PointF(x + barW / 2 - 25, baseY + 8));
            }

            DrawBar(60, _lastBulkSummary.PositiveCount,
                GreenColor, Color.FromArgb(20, 180, 70), "Positive");
            DrawBar(200, _lastBulkSummary.NegativeCount,
                RedColor, Color.FromArgb(180, 40, 40), "Negative");

            if (_lastBulkSummary.SkippedCount > 0)
                DrawBar(340, _lastBulkSummary.SkippedCount,
                    Color.Orange, Color.FromArgb(180, 120, 0), "Skipped");

            using var bp = new Pen(MutedColor, 1.5f);
            g.DrawLine(bp, 40, baseY, 430, baseY);
        }

        private void DrawNoData(Graphics g, Rectangle rect)
        {
            using var f = new Font("Segoe UI", 12f);
            g.DrawString("No data yet.\nRun Bulk Analysis first.",
                f, new SolidBrush(MutedColor),
                new RectangleF(0, 0, rect.Width, rect.Height),
                new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                });
        }

        // ============================================================
        // EXPORT
        // ============================================================
        private async void BtnExportCsv_Click(object? sender, EventArgs e)
        {
            if (_lastBulkSummary == null)
            {
                MessageBox.Show("Run Bulk Analysis first.", "No Data",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            using var dlg = new SaveFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv",
                FileName = $"sentiment_results_{DateTime.Now:yyyyMMdd_HHmm}.csv"
            };
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                await _sentimentService.ExportToCsvAsync(_lastBulkSummary, dlg.FileName);
                rtbLog.AppendText($"[{DateTime.Now:HH:mm:ss}] ✅ CSV exported: {dlg.FileName}\n");
                MessageBox.Show("CSV exported successfully!", "Done",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnExportPdf_Click(object? sender, EventArgs e)
        {
            if (_lastBulkSummary == null)
            {
                MessageBox.Show("Run Bulk Analysis first.", "No Data",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            using var dlg = new SaveFileDialog
            {
                Filter = "PDF Files (*.pdf)|*.pdf",
                FileName = $"sentiment_report_{DateTime.Now:yyyyMMdd_HHmm}.pdf"
            };
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    PdfExporter.Export(_lastBulkSummary, dlg.FileName);
                    rtbLog.AppendText($"[{DateTime.Now:HH:mm:ss}] ✅ PDF exported: {dlg.FileName}\n");
                    MessageBox.Show("PDF report exported successfully!", "Done",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    rtbLog.AppendText($"[{DateTime.Now:HH:mm:ss}] ❌ PDF Error: {ex.Message}\n");
                    MessageBox.Show($"PDF export failed:\n{ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // ============================================================
        // THEME TOGGLE - COMPLETE FIX
        // ============================================================
        private void ToggleTheme()
        {
            _isDarkMode = !_isDarkMode;
            btnTheme.Text = _isDarkMode ? "☀  Light Mode" : "🌙  Dark Mode";
            ApplyThemeToAll();
        }

        private void ApplyThemeToAll()
        {
            // Form
            BackColor = BgColor;

            // Header
            pnlHeader.BackColor = PanelColor;
            lblTitle.ForeColor = TextColor;
            lblSub.ForeColor = MutedColor;
            btnTheme.BackColor = CardColor;
            btnTheme.ForeColor = TextColor;
            btnTheme.FlatAppearance.BorderColor = BorderColor;

            // Status
            pnlStatus.BackColor = CardColor;
            lblStatus.ForeColor = _sentimentService.IsReady ? GreenColor : AccentColor;

            // Tab pages
            foreach (TabPage tp in tcMain.TabPages)
            {
                tp.BackColor = BgColor;
                ApplyThemeToControls(tp.Controls);
            }

            // Tab1 specific
            if (pnl1 != null) pnl1.BackColor = BgColor;
            if (rtbInput != null) { rtbInput.BackColor = CardColor; rtbInput.ForeColor = TextColor; }
            if (pnlResult != null) pnlResult.BackColor = CardColor;
            if (btnClear != null)
            {
                btnClear.BackColor = CardColor;
                btnClear.ForeColor = TextColor;
                btnClear.FlatAppearance.BorderColor = BorderColor;
            }

            // Tab2 specific
            if (pnl2 != null) pnl2.BackColor = BgColor;
            if (lblFile != null) lblFile.ForeColor = MutedColor;
            if (pnlStats != null) pnlStats.BackColor = CardColor;
            if (btnBrowse != null)
            {
                btnBrowse.BackColor = CardColor;
                btnBrowse.ForeColor = TextColor;
                btnBrowse.FlatAppearance.BorderColor = AccentColor;
            }

            // Grid
            if (dgvResults != null)
            {
                dgvResults.BackgroundColor = CardColor;
                dgvResults.GridColor = BorderColor;
                dgvResults.DefaultCellStyle.BackColor = CardColor;
                dgvResults.DefaultCellStyle.ForeColor = TextColor;
                dgvResults.AlternatingRowsDefaultCellStyle.BackColor =
                    _isDarkMode ? Color.FromArgb(30, 30, 50) : Color.FromArgb(242, 244, 255);
                dgvResults.AlternatingRowsDefaultCellStyle.ForeColor = TextColor;
                dgvResults.ColumnHeadersDefaultCellStyle.BackColor =
                    _isDarkMode ? Color.FromArgb(50, 50, 80) : Color.FromArgb(180, 185, 220);
                dgvResults.ColumnHeadersDefaultCellStyle.ForeColor =
                    _isDarkMode ? Color.White : Color.FromArgb(20, 20, 50);
                dgvResults.Invalidate();
            }

            // Charts
            if (pnlPie != null) { pnlPie.BackColor = CardColor; pnlPie.Invalidate(); }
            if (pnlBar != null) { pnlBar.BackColor = CardColor; pnlBar.Invalidate(); }

            // Export
            if (rtbLog != null) { rtbLog.BackColor = CardColor; rtbLog.ForeColor = GreenColor; }

            Invalidate(true);
            Update();
        }

        private void ApplyThemeToControls(Control.ControlCollection controls)
        {
            foreach (Control c in controls)
            {
                if (c is Panel p) p.BackColor = BgColor;
                if (c is Label l && l != lblSentiment && l != lblPos
                    && l != lblNeg && l != lblMood)
                {
                    if (l.ForeColor != GreenColor && l.ForeColor != RedColor
                        && l.ForeColor != Color.Orange && l.ForeColor != AccentColor)
                        l.ForeColor = TextColor;
                }
                if (c.HasChildren)
                    ApplyThemeToControls(c.Controls);
            }
        }

        // ============================================================
        // HELPERS
        // ============================================================
        private Label MkLabel(string text, int x, int y,
            float size = 9.5f, FontStyle style = FontStyle.Regular)
        {
            return new Label
            {
                Text = text,
                Location = new Point(x, y),
                Font = new Font("Segoe UI", size, style),
                ForeColor = TextColor,
                AutoSize = true
            };
        }

        private Button MkBtn(string text, int x, int y, Color bg)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                FlatStyle = FlatStyle.Flat,
                BackColor = bg,
                ForeColor = bg == AccentColor || bg == GreenColor
                    ? Color.White : TextColor,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor =
                ControlPaint.Light(bg, 0.15f);
            return btn;
        }

        private Label MkStatBox(string label, string value, int x, Panel parent)
        {
            var lbl = new Label
            {
                Text = label,
                Location = new Point(x, 10),
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = MutedColor,
                AutoSize = true
            };
            var val = new Label
            {
                Text = value,
                Location = new Point(x, 30),
                Font = new Font("Segoe UI", 17f, FontStyle.Bold),
                ForeColor = TextColor,
                AutoSize = true
            };
            parent.Controls.AddRange(new Control[] { lbl, val });
            return val;
        }
    }

    public static class GraphicsExtensions
    {
        public static void DrawRoundedRectangle(this Graphics g,
            Pen pen, Rectangle rect, int radius)
        {
            using var path = new GraphicsPath();
            path.AddArc(rect.X, rect.Y, radius * 2, radius * 2, 180, 90);
            path.AddArc(rect.Right - radius * 2, rect.Y, radius * 2, radius * 2, 270, 90);
            path.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2,
                radius * 2, radius * 2, 0, 90);
            path.AddArc(rect.X, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseFigure();
            g.DrawPath(pen, path);
        }
    }
}
