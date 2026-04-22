# Tech Stack - D2ViewerEditor

## Frontend

- Angular 20 (standalone components + signals)
- TypeScript 5.8
- RxJS 7.8
- SCSS
- Vitest do testów jednostkowych

## Backend

- ASP.NET Core Web API (.NET 8)
- Clean Architecture (Api/Application/Domain/Infrastructure)
- MediatR (CQRS)
- FluentValidation
- Swagger (Swashbuckle)

## Document Processing

- DocumentFormat.OpenXml do generowania/parsowania DOCX
- HtmlAgilityPack do obsługi HTML

## Barcode and Imaging

- ZXing.Net do kodowania barcode/QR
- SkiaSharp do renderowania obrazów

## Storage

- PostgreSQL 16 do metadanych i historii wersji
- Google Cloud Storage client (lub fake-gcs w DEV)

## Testing

- .NET unit tests: NUnit + FluentAssertions + NSubstitute
- E2E: Python + Playwright + pytest-bdd
- Performance: Locust

## Infrastructure / Runtime

- Podman + podman-compose
- Nginx do serwowania frontendu w setupie kontenerowym
- Swagger UI do manualnego testowania API

