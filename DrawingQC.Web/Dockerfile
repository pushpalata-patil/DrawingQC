# Build stage. Context is the repo ROOT so the web project can pull in the shared
# QcEngine.cs that lives in ..\DrawingQC.UI (referenced via <Compile Include> in the csproj).
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src
COPY . .
RUN dotnet publish DrawingQC.Web/DrawingQC.Web.csproj -c Release -o /app /p:UseAppHost=false

# Local Agent: the Windows companion that drives the user's own AutoCAD / Plant 3D and Word.
# Cross-published here as a self-contained single .exe (no .NET install needed on the PC) and
# zipped into wwwroot so the site can offer it for download at /agent/DrawingQC.Agent.zip.
RUN apt-get update && apt-get install -y --no-install-recommends zip && rm -rf /var/lib/apt/lists/* \
 && dotnet publish DrawingQC.Agent/DrawingQC.Agent.csproj -c Release -r win-x64 --self-contained true \
      -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=none -o /agent \
 && cp DrawingQC.Agent/README.txt /agent/ \
 && mkdir -p /app/wwwroot/agent \
 && cd /agent && zip -r -q /app/wwwroot/agent/DrawingQC.Agent.zip .

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:10.0
WORKDIR /app
COPY --from=build /app ./
ENV ASPNETCORE_ENVIRONMENT=Production
# The app binds to the PORT env var when hosted (Render sets it automatically).
EXPOSE 8080
ENTRYPOINT ["dotnet", "DrawingQC.Web.dll"]
