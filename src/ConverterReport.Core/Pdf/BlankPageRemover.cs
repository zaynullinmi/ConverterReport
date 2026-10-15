using System.Text;
using System.Text.RegularExpressions;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;

namespace ConverterReport.Core.Pdf;

/// <summary>Удаляет из PDF пустые страницы (без текста, графики и изображений).</summary>
public static class BlankPageRemover
{
    // Операторы, которые что-то выводят на страницу: текст, изображения/формы, закраска и обводка контуров.
    private static readonly HashSet<string> PaintOperators = new(StringComparer.Ordinal)
    {
        "Tj", "TJ", "'", "\"", "Do", "BI", "sh", "S", "s", "f", "F", "f*", "B", "B*", "b", "b*",
    };

    /// <summary>Возвращает число удаленных страниц. Последняя оставшаяся страница не удаляется.</summary>
    public static int RemoveBlankPages(PdfDocument document)
    {
        var removed = 0;
        for (var i = document.PageCount - 1; i >= 0 && document.PageCount > 1; i--)
        {
            if (!IsBlank(document.Pages[i]))
                continue;
            document.Pages.RemoveAt(i);
            removed++;
        }

        return removed;
    }

    public static bool IsBlank(PdfPage page)
    {
        var content = new StringBuilder();
        foreach (var item in page.Contents.Elements)
        {
            var dict = item switch
            {
                PdfReference reference => reference.Value as PdfDictionary,
                PdfDictionary d => d,
                _ => null,
            };
            if (dict?.Stream is { } stream)
                content.Append(Encoding.Latin1.GetString(stream.UnfilteredValue)).Append('\n');
        }

        if (content.Length == 0)
            return true;

        // Строковые литералы (...) не учитываем — в них могут встречаться буквы, похожие на операторы.
        var withoutStrings = Regex.Replace(content.ToString(), @"\((?:\\.|[^\\)])*\)", " ");
        return !Regex.Split(withoutStrings, @"[\s\[\]<>]+").Any(PaintOperators.Contains);
    }
}
