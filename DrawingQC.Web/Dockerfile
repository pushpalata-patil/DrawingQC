# Build stage
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src
COPY DrawingQC.Web.csproj ./
RUN dotnet restore
COPY . ./
RUN dotnet publish DrawingQC.Web.csproj -c Release -o /app /p:UseAppHost=false

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:10.0
WORKDIR /app
COPY --from=build /app ./
ENV ASPNETCORE_ENVIRONMENT=Production
# The app binds to the port given by the PORT env var (Render sets this automatically).
EXPOSE 8080
ENTRYPOINT ["dotnet", "DrawingQC.Web.dll"]
