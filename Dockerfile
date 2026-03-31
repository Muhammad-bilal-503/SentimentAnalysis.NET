# Base image - .NET 9.0 SDK
FROM mcr.microsoft.com/dotnet/sdk:9.0 AS build

# Working directory set karo
WORKDIR /app

# Project file copy karo
COPY *.csproj ./

# Dependencies restore karo
RUN dotnet restore

# Baaki sab files copy karo
COPY . .

# Application build karo
RUN dotnet build -c Release

# Run karo
CMD ["dotnet", "run", "--project", "SentimentAnalyzerPro.csproj"]