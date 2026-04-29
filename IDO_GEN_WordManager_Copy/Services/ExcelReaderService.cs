using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IDO_GEN_WordManager.Services
{
    public class ExcelReaderService
    {
        public Dictionary<int, (string Header, List<string> Values)> LoadColumns(string filePath)
        {
            var result = new Dictionary<int, (string Header, List<string> Values)>();
            if (!File.Exists(filePath)) return result;

            using var doc = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = doc.WorkbookPart;
            if (workbookPart == null) return result;

            var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;

            var sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault();
            if (sheet?.Id?.Value == null) return result;

            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
            var rows = worksheetPart.Worksheet.Descendants<Row>().ToList();
            if (rows.Count == 0) return result;

            // Primera fila = encabezados de columna
            foreach (var cell in rows[0].Elements<Cell>())
            {
                var colIdx = GetColumnIndex(cell.CellReference?.Value ?? "A1");
                var header = GetCellValue(cell, sharedStrings);
                if (!string.IsNullOrWhiteSpace(header))
                    result[colIdx] = (header.Trim(), new List<string>());
            }

            // Filas restantes = valores
            for (int i = 1; i < rows.Count; i++)
            {
                foreach (var cell in rows[i].Elements<Cell>())
                {
                    var colIdx = GetColumnIndex(cell.CellReference?.Value ?? "A1");
                    if (result.ContainsKey(colIdx))
                    {
                        var value = GetCellValue(cell, sharedStrings);
                        if (!string.IsNullOrWhiteSpace(value))
                            result[colIdx].Values.Add(value.Trim());
                    }
                }
            }

            // Quitar columnas sin valores
            foreach (var k in result.Keys.Where(k => result[k].Values.Count == 0).ToList())
                result.Remove(k);

            return result;
        }

        private static string GetCellValue(Cell cell, SharedStringTable? sharedStrings)
        {
            var value = cell.InnerText;
            if (cell.DataType?.Value == CellValues.SharedString && sharedStrings != null
                && int.TryParse(value, out int idx))
            {
                return sharedStrings.ElementAt(idx).InnerText;
            }
            return value ?? string.Empty;
        }

        private static int GetColumnIndex(string cellReference)
        {
            var letters = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
            int index = 0;
            foreach (char c in letters.ToUpperInvariant())
                index = index * 26 + (c - 'A' + 1);
            return index;
        }
    }
}
