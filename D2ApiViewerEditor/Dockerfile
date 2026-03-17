# ── Stage 1: build ──────────────────────────────────────────────────────────
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Kopiuj pliki projektów i przywróć zależności (cache warstw)
COPY D2ViewerEditor.sln ./
COPY D2ViewerEditor.Api/D2ViewerEditor.Api.csproj                         D2ViewerEditor.Api/
COPY D2ViewerEditor.Application/D2ViewerEditor.Application.csproj         D2ViewerEditor.Application/
COPY D2ViewerEditor.Domain/D2ViewerEditor.Domain.csproj                   D2ViewerEditor.Domain/
COPY D2ViewerEditor.Infrastructure/D2ViewerEditor.Infrastructure.csproj   D2ViewerEditor.Infrastructure/

RUN dotnet restore D2ViewerEditor.Api/D2ViewerEditor.Api.csproj

# Kopiuj resztę kodu i opublikuj
COPY . .
RUN dotnet publish D2ViewerEditor.Api/D2ViewerEditor.Api.csproj \
    -c Release \
    -o /app/publish \
    --no-restore

# ── Stage 2: runtime ─────────────────────────────────────────────────────────
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app

# Nie uruchamiaj jako root
RUN addgroup --system appgroup && adduser --system --ingroup appgroup appuser
USER appuser

COPY --from=build /app/publish .

ENV ASPNETCORE_URLS=http://+:8080
ENV ASPNETCORE_ENVIRONMENT=DEV

EXPOSE 8080

ENTRYPOINT ["dotnet", "D2ViewerEditor.Api.dll"]
