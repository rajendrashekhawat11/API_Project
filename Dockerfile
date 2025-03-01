# Use official .NET 8.0 runtime as base image
FROM mcr.microsoft.com/dotnet/sdk:8.0

# Set working directory inside container
WORKDIR /app

# Copy everything to the container
COPY . .

# Restore and build the project
RUN dotnet restore
RUN dotnet build --configuration Release

# Command to run the application
CMD ["dotnet", "run"]
