using System.Xml;
using System.Xml.Linq;
using D2ViewerEditor.Domain.Interfaces;

namespace D2ViewerEditor.Infrastructure.Services.StructureInspection;

/// <summary>
/// Leniwy odczyt surowego XML: fragment wskazanego elementu oraz numer linii, od której zaczyna się
/// on w części źródłowej. Element jest wskazywany ścieżką indeksów dzieci, więc odczyt nie zależy
/// od prefiksów namespace ani od wcześniej zbudowanego drzewa obiektów. Parsowanie części należy do
/// wołającego (inspector cache'uje sparsowane dokumenty per pakiet).
/// </summary>
public sealed class OoxmlFragmentReader
{
    public OoxmlFragment? Read(XDocument document, IReadOnlyList<int> nodePath)
    {
        var node = Resolve(document, nodePath);

        return node is null ? null : new OoxmlFragment(SerializeStandalone(node), GetLineNumber(node));
    }

    public int? FindLine(XDocument document, IReadOnlyList<int> nodePath) =>
        GetLineNumber(Resolve(document, nodePath));

    private static XElement? Resolve(XDocument document, IReadOnlyList<int> nodePath)
    {
        var current = document.Root;

        foreach (var childIndex in nodePath)
        {
            if (current is null || childIndex < 0)
            {
                return null;
            }

            current = current.Elements().ElementAtOrDefault(childIndex);
        }

        return current;
    }

    private static int? GetLineNumber(XElement? node) =>
        node is IXmlLineInfo lineInfo && lineInfo.HasLineInfo() ? lineInfo.LineNumber : null;

    /// <summary>
    /// Fragment samodzielny: element z deklaracjami namespace obowiązującymi w jego miejscu
    /// w dokumencie. Bez tego prefiksy w podglądzie (np. <c>wp:</c>, <c>a:</c>) byłyby nierozwiązane.
    /// </summary>
    private static string SerializeStandalone(XElement element)
    {
        var clone = new XElement(element);
        var content = clone.Nodes().ToArray();
        clone.RemoveNodes();

        foreach (var declaration in element.AncestorsAndSelf().Reverse()
                     .SelectMany(ancestor => ancestor.Attributes().Where(attribute => attribute.IsNamespaceDeclaration)))
        {
            clone.Attribute(declaration.Name)?.Remove();
            clone.Add(new XAttribute(declaration));
        }

        clone.Add(content);

        return clone.ToString(SaveOptions.None);
    }
}
