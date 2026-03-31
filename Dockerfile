FROM mcr.microsoft.com/dotnet/sdk:9.0
WORKDIR /app
COPY . .
RUN dotnet build /p:EnableWindowsTargeting=true
CMD ["dotnet", "run"]s