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
            var toHide = new HashSet<int>(inactive.Where(h => h.Action == HeadingAction.Hide).Select(h => h.ParagraphIndex));

            using var doc = WordprocessingDocument.Open(destPath, true);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return;

            var paragraphs = body.Elements<Paragraph>().ToList();

            // Ocultar: aplicar marcas para vistas paginadas y web.
            foreach (var idx in toHide)
            {
                if (idx >= paragraphs.Count) continue;
                var para = paragraphs[idx];
                var ppr = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
                var pmr = ppr.ParagraphMarkRunProperties ?? ppr.AppendChild(new ParagraphMarkRunProperties());

                if (pmr.GetFirstChild<Vanish>() == null)
                    pmr.AppendChild(new Vanish());

                if (pmr.GetFirstChild<WebHidden>() == null)
                    pmr.AppendChild(new WebHidden());

                foreach (var run in para.Descendants<Run>())
                {
                    var rpr = run.RunProperties ?? run.PrependChild(new RunProperties());
                    rpr.Vanish = new Vanish();
                    rpr.WebHidden = new WebHidden();
                }
            }

            // Eliminar: quitar el parrafo por completo
            foreach (var idx in toDelete.OrderByDescending(i => i))
            {
                if (idx < paragraphs.Count)
                    paragraphs[idx].Remove();
            }

            doc.MainDocumentPart!.Document.Save();
        }

        /// <summary>
        /// Elimina las marcas de oculto de todos los runs del documento,
        /// haciendo visible todo el texto que fue ocultado programaticamente.
        /// </summary>
        public int UnhideAllVanish(string filePath)
        {
            int count = 0;
            using var doc = WordprocessingDocument.Open(filePath, true);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return count;

            foreach (var para in body.Descendants<Paragraph>())
            {
                var pmr = para.ParagraphProperties?.ParagraphMarkRunProperties;
                if (pmr == null) continue;

                bool changed = false;
                var paragraphVanish = pmr.GetFirstChild<Vanish>();
                if (paragraphVanish != null)
                {
                    paragraphVanish.Remove();
                    changed = true;
                }

                var paragraphWebHidden = pmr.GetFirstChild<WebHidden>();
                if (paragraphWebHidden != null)
                {
                    paragraphWebHidden.Remove();
                    changed = true;
                }

                if (!changed) continue;

                count++;
                if (!pmr.HasChildren)
                    pmr.Remove();
            }

            foreach (var run in body.Descendants<Run>())
            {
                var rpr = run.RunProperties;
                if (rpr == null) continue;

                bool changed = false;
                if (rpr.Vanish != null)
                {
                    rpr.Vanish.Remove();
                    changed = true;
                }

                if (rpr.WebHidden != null)
                {
                    rpr.WebHidden.Remove();
                    changed = true;
                }

                if (!changed) continue;

                count++;
                if (!rpr.HasChildren)
                    rpr.Remove();
            }

            doc.MainDocumentPart!.Document.Save();
            return count;
        }
    }
}
