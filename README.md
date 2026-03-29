# 🧠 Sentiment Analyzer Pro

A multi-language AI-powered sentiment analysis desktop application built with **C# .NET 6 + ML.NET + Windows Forms**.

---

## 📋 Features

- ✅ **Single Text Analysis** — type any comment and get instant sentiment
- ✅ **Bulk CSV Analysis** — upload a CSV file with multiple comments
- ✅ **Multi-Language Support** — auto-detect & translate via LibreTranslate
- ✅ **Pie Chart & Bar Chart** — visual results
- ✅ **Export to CSV** — save results
- ✅ **Export PDF Report** — professional report
- ✅ **Dark / Light Mode** — toggle theme
- ✅ **Offline Fallback** — English comments work without internet

---

## 🚀 Setup Instructions

### Step 1 — Prerequisites
- Visual Studio 2022
- .NET 6.0 SDK

### Step 2 — Clone or Open Project
Open `SentimentAnalyzerPro.sln` in Visual Studio 2022.

### Step 3 — Add Dataset
1. Download `IMDB Dataset.csv` from Kaggle
2. Place it inside the `Data/` folder in the project
3. Right-click the file in Solution Explorer → Properties → **Copy to Output: PreserveNewest**

### Step 4 — Restore NuGet Packages
In Visual Studio:
```
Tools → NuGet Package Manager → Restore Packages
```

### Step 5 — Run
Press **F5** to build and run.

On first launch, the model will train automatically (~1-2 minutes).
After that, `sentiment_model.zip` is saved and loads instantly.

---

## 📂 Project Structure

```
SentimentAnalyzerPro/
├── Models/
│   └── ReviewData.cs          # Data classes
├── Services/
│   ├── MLService.cs           # ML.NET training & prediction
│   ├── TranslationService.cs  # LibreTranslate API
│   ├── SentimentService.cs    # Main orchestrator
│   └── PdfExporter.cs         # PDF export
├── Forms/
│   └── MainForm.cs            # Windows Forms UI
├── Data/
│   └── IMDB Dataset.csv       # Training dataset (add manually)
└── Program.cs
```

---

## 🌍 Supported Languages

Urdu 🇵🇰 | Hindi 🇮🇳 | Arabic 🇸🇦 | French 🇫🇷 | German 🇩🇪 | Spanish 🇪🇸 | Chinese 🇨🇳 | and more!

---

## 📌 Notes

- Internet required for non-English translation
- English comments always work offline
- Non-English comments without internet are marked as "Skipped"
- Model trains once and saves as `sentiment_model.zip`

---

## 👨‍💻 Developer

**Muhammad Bilal** | Roll No: 232201100  
Khan Institute of Computer Science  
Submitted to: Sir Uzair Hassan
