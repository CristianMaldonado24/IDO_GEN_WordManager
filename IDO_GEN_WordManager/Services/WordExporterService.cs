using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using IDO_GEN_WordManager.Models;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IDO_GEN_WordManager.Services
{
    public class WordExporterService
    {
        public void ExportFiltered(string sourcePath, string destPath, IEnumerable<DocumentHeading> headings)
        {
            File.Copy(sourcePath, destPath, overwrite: true);

            var inactive = headings.Where(h => !h.IsVisible).ToList();
            if (inactive.Count == 0) return;

            var toDelete = new HashSet<int>(inactive.Where(h => h.Action == HeadingAction.Delete).Select(h => h.ParagraphIndex));
            var toHide   = new HashSet<int>(inactive.Where(h => h.Action == HeadingAction.Hide).Select(h => h.ParagraphIndex));

            using var doc = WordprocessingDocument.Open(destPath, true);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return;

            var paragraphs = body.Elements<Paragraph>().ToList();

            // Ocultar: hacer el texto blanco (invisible pero conserva espacio en el layout)
            foreach (var idx in toHide)
            {
                if (idx >= paragraphs.Count) continue;
                var para = paragraphs[idx];
                foreach (var run in para.Descendants<Run>())
                {
                    var rpr = run.RunProperties ?? run.PrependChild(new RunProperties());
                    // Color blanco = invisible
                    rpr.Color = new Color { Val = "FFFFFF" };
                    rpr.FontSize = new FontSize { Val = "2" };     // 1pt, ocupa mínimo espacio
                }
            }

            // Eliminar: quitar el párrafo por completo
            foreach (var idx in toDelete.OrderByDescending(i => i))
            {
                if (idx < paragraphs.Count)
                    paragraphs[idx].Remove();
            }

            doc.MainDocumentPart!.Document.Save();
        }
    }
}
