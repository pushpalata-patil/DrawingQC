# Build stage. Context is the repo ROOT so the web project can pull in the shared
# QcEngine.cs that lives in ..\DrawingQC.UI (referenced via <Compile Include> in the csproj).
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src
COPY . .
RUN dotnet publish DrawingQC.Web/DrawingQC.Web.csproj -c Release -o /app /p:UseAppHost=false

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:10.0
WORKDIR /app
COPY --from=build /app ./
ENV ASPNETCORE_ENVIRONMENT=Production
# The app binds to the PORT env var when hosted (Render sets it automatically).
EXPOSE 8080
ENTRYPOINT ["dotnet", "DrawingQC.Web.dll"]
