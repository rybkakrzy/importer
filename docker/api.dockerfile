FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

COPY D2ApiViewerEditor/D2ViewerEditor.sln ./
COPY D2ApiViewerEditor/D2ViewerEditor.Api/D2ViewerEditor.Api.csproj                         D2ViewerEditor.Api/
COPY D2ApiViewerEditor/D2ViewerEditor.Application/D2ViewerEditor.Application.csproj         D2ViewerEditor.Application/
COPY D2ApiViewerEditor/D2ViewerEditor.Domain/D2ViewerEditor.Domain.csproj                   D2ViewerEditor.Domain/
COPY D2ApiViewerEditor/D2ViewerEditor.Infrastructure/D2ViewerEditor.Infrastructure.csproj   D2ViewerEditor.Infrastructure/

RUN dotnet restore D2ViewerEditor.Api/D2ViewerEditor.Api.csproj

COPY D2ApiViewerEditor/ .
RUN dotnet publish D2ViewerEditor.Api/D2ViewerEditor.Api.csproj \
    -c Release \
    -o /app/publish \
    --no-restore

FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app

# LibreOffice headless — używany do rasteryzacji EMF/WMF (vector Windows metafiles)
# osadzonych w dokumentach DOCX. Bez tego obrazki w formacie metafile nie wyświetlą
# się w przeglądarce. Instalujemy tylko core + draw (minimalny zestaw do --convert-to).
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-core \
        libreoffice-draw \
        fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

# Wskaż binarkę LibreOffice dla aplikacji (omija narzut probe'owania PATH).
ENV SOFFICE_BIN=/usr/bin/soffice

RUN addgroup --system appgroup && adduser --system --ingroup appgroup appuser
USER appuser

COPY --from=build /app/publish .

ENV ASPNETCORE_URLS=http://+:8080
ENV ASPNETCORE_ENVIRONMENT=DEV

EXPOSE 8080

ENTRYPOINT ["dotnet", "D2ViewerEditor.Api.dll"]
