using D2Tools.Domain.Common;
using D2Tools.Domain.Models;
using MediatR;

namespace D2Tools.Application.Features.Documents.Queries.GetTemplate;

public class GetTemplateQueryHandler : IRequestHandler<GetTemplateQuery, Result<DocumentContent>>
{
    public Task<Result<DocumentContent>> Handle(GetTemplateQuery request, CancellationToken cancellationToken)
    {
        var html = request.TemplateId switch
        {
            "blank" => "<div class=\"document-content\"><p>&nbsp;</p></div>",

            "letter" => @"<div class=""document-content"">
                <p style=""text-align:right;"">Miejscowość, data</p>
                <p>&nbsp;</p>
                <p><strong>Nadawca:</strong></p>
                <p>Imię i Nazwisko</p>
                <p>Adres</p>
                <p>&nbsp;</p>
                <p><strong>Odbiorca:</strong></p>
                <p>Imię i Nazwisko</p>
                <p>Adres</p>
                <p>&nbsp;</p>
                <p style=""text-align:center;""><strong>Temat listu</strong></p>
                <p>&nbsp;</p>
                <p>Szanowni Państwo,</p>
                <p>&nbsp;</p>
                <p>Treść listu...</p>
                <p>&nbsp;</p>
                <p>Z poważaniem,</p>
                <p>Podpis</p>
            </div>",

            "report" => @"<div class=""document-content"">
                <h1 style=""text-align:center;"">Tytuł Raportu</h1>
                <p style=""text-align:center;"">Autor: Imię Nazwisko</p>
                <p style=""text-align:center;"">Data: " + DateTime.Now.ToString("dd.MM.yyyy") + @"</p>
                <p>&nbsp;</p>
                <h2>1. Wprowadzenie</h2>
                <p>Treść wprowadzenia...</p>
                <p>&nbsp;</p>
                <h2>2. Analiza</h2>
                <p>Treść analizy...</p>
                <p>&nbsp;</p>
                <h2>3. Wnioski</h2>
                <p>Treść wniosków...</p>
                <p>&nbsp;</p>
                <h2>4. Podsumowanie</h2>
                <p>Treść podsumowania...</p>
            </div>",

            "cv" => @"<div class=""document-content"">
                <h1 style=""text-align:center;"">Imię i Nazwisko</h1>
                <p style=""text-align:center;"">email@example.com | +48 123 456 789 | Miasto</p>
                <p>&nbsp;</p>
                <h2>Doświadczenie zawodowe</h2>
                <p><strong>Stanowisko</strong> | Firma | Okres</p>
                <p>Opis obowiązków...</p>
                <p>&nbsp;</p>
                <h2>Wykształcenie</h2>
                <p><strong>Kierunek</strong> | Uczelnia | Okres</p>
                <p>&nbsp;</p>
                <h2>Umiejętności</h2>
                <p>• Umiejętność 1</p>
                <p>• Umiejętność 2</p>
                <p>• Umiejętność 3</p>
                <p>&nbsp;</p>
                <h2>Języki</h2>
                <p>• Polski - ojczysty</p>
                <p>• Angielski - zaawansowany</p>
            </div>",

            _ => "<div class=\"document-content\"><p>&nbsp;</p></div>"
        };

        var title = request.TemplateId switch
        {
            "letter" => "Nowy list",
            "report" => "Nowy raport",
            "cv" => "Nowe CV",
            _ => "Nowy dokument"
        };

        var content = new DocumentContent
        {
            Html = html,
            Metadata = new DocumentMetadata
            {
                Title = title,
                Created = DateTime.Now,
                Modified = DateTime.Now
            },
            Images = new List<DocumentImage>()
        };

        return Task.FromResult(Result<DocumentContent>.Success(content));
    }
}
