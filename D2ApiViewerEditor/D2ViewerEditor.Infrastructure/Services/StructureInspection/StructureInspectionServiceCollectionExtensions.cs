using D2ViewerEditor.Application.Features.StructureInspection.Common;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Infrastructure.Services.StructureInspection.Analyzers;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace D2ViewerEditor.Infrastructure.Services.StructureInspection;

/// <summary>
/// Rejestracja warstwy diagnostyki struktury OOXML. Wszystkie komponenty analizy są bezstanowe
/// (stan trzyma wyłącznie magazyn analiz), więc mogą być singletonami. Dodanie nowego analizatora
/// = jedna linia rejestracji <see cref="IStructureAnalyzer"/>; kolejność rejestracji nie zmienia
/// wyniku, bo analizatory są niezależne.
/// </summary>
public static class StructureInspectionServiceCollectionExtensions
{
    public static IServiceCollection AddStructureInspection(
        this IServiceCollection services,
        IConfiguration configuration)
    {
        services.Configure<StructureInspectionOptions>(
            configuration.GetSection(StructureInspectionOptions.SectionName));
        services.Configure<EditorCompatibilityOptions>(
            configuration.GetSection(EditorCompatibilityOptions.SectionName));
        services.Configure<StructureInspectionRetentionOptions>(
            configuration.GetSection(StructureInspectionOptions.SectionName));

        services.AddSingleton<SafeOoxmlXmlLoader>();
        services.AddSingleton<OoxmlPackageReader>();
        services.AddSingleton<OpcPackageAnalyzer>();
        services.AddSingleton<OoxmlElementClassifier>();
        services.AddSingleton<OoxmlElementIndexer>();
        services.AddSingleton<OoxmlFragmentReader>();
        services.AddSingleton<OpenXmlSchemaValidatorRunner>();
        services.AddSingleton<SchemaIssueMapper>();
        services.AddSingleton<SectionHeaderFooterAnalyzer>();

        services.AddSingleton<IStructureAnalyzer, EffectiveFormattingAnalyzer>();
        services.AddSingleton<IStructureAnalyzer, NumberingAnalyzer>();
        services.AddSingleton<IStructureAnalyzer, TableAnalyzer>();
        services.AddSingleton<IStructureAnalyzer, DrawingAnalyzer>();
        services.AddSingleton<IStructureAnalyzer, SectionLayoutAnalyzer>();
        services.AddSingleton<IStructureAnalyzer, FieldAnalyzer>();
        services.AddSingleton<IStructureAnalyzer, ReferenceAnalyzer>();
        services.AddSingleton<IStructureAnalyzer, ContentControlAnalyzer>();
        services.AddSingleton<IStructureAnalyzer, RevisionAnalyzer>();
        services.AddSingleton<IStructureAnalyzer, MarkupCompatibilityAnalyzer>();
        services.AddSingleton<IStructureAnalyzer, RunAndParagraphFeatureAnalyzer>();

        // Profil zgodności edytora działa NA KOŃCU: czyta kategorie i atrybuty ustalone wcześniej.
        services.AddSingleton<IStructureAnalyzer, EditorCompatibilityAnalyzer>();

        services.AddSingleton<IDocumentStructureInspector, DocumentStructureInspector>();
        services.AddSingleton<IDocumentStructureInspectionStore, InMemoryDocumentStructureInspectionStore>();

        return services;
    }
}
