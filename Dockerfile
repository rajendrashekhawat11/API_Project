# Use official .NET 8.0 runtime as base image
FROM mcr.microsoft.com/dotnet/sdk:8.0

# Set working directory inside container
WORKDIR /app

# Copy project file and restore as distinct layers
COPY *.csproj ./
RUN dotnet restore

# Copy everything else and build
COPY . .
RUN dotnet add package Google.Apis.Sheets.v4
RUN dotnet add package Google.Apis.Auth
RUN dotnet add package Google.Apis.Core
RUN dotnet add package Google.Apis
RUN dotnet build --configuration Release --no-restore

# Command to run the application
CMD ["dotnet", "run"]
