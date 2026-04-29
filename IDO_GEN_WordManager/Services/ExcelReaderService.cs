using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace IDO_GEN_WordManager.Services
{
    public class ExcelSheetData
    {
        public string SheetName { get; set; } = string.Empty;
        public Dictionary<int, (string Header, List<string> Values)> Columns { get; set; } = new();
    }

    public class ExcelReaderService
    {
        public List<ExcelSheetData> LoadWorkbook(string filePath)
        {
            var workbookData = new List<ExcelSheetData>();
            if (!File.Exists(filePath)) return workbookData;

            using var doc = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = doc.WorkbookPart;
            if (workbookPart == null) return workbookData;

            var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
            var styles = workbookPart.WorkbookStylesPart?.Stylesheet;
            var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();

            foreach (var sheet in sheets)
            {
                if (sheet.Id?.Value == null) continue;

                var result = new Dictionary<int, (string Header, List<string> Values)>();
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
                var rows = worksheetPart.Worksheet.Descendants<Row>().ToList();
                if (rows.Count == 0) continue;

                foreach (var cell in rows[0].Elements<Cell>())
                {
                    var colIdx = GetColumnIndex(cell.CellReference?.Value ?? "A1");
                    var header = GetCellValue(cell, sharedStrings, styles);
                    if (!string.IsNullOrWhiteSpace(header))
                        result[colIdx] = (header.Trim(), new List<string>());
                }

                for (int i = 1; i < rows.Count; i++)
                {
                    foreach (var cell in rows[i].Elements<Cell>())
                    {
                        var colIdx = GetColumnIndex(cell.CellReference?.Value ?? "A1");
                        if (result.ContainsKey(colIdx))
                        {
                            var value = GetCellValue(cell, sharedStrings, styles);
                            if (!string.IsNullOrWhiteSpace(value))
                                result[colIdx].Values.Add(value.Trim());
                        }
                    }
                }

                foreach (var k in result.Keys.Where(k => result[k].Values.Count == 0).ToList())
                    result.Remove(k);

                if (result.Count == 0) continue;

                workbookData.Add(new ExcelSheetData
                {
                    SheetName = sheet.Name?.Value ?? "Hoja",
                    Columns = result
                });
            }

            return workbookData;
        }

        private static string GetCellValue(Cell cell, SharedStringTable? sharedStrings, Stylesheet? styles)
        {
            var value = cell.CellValue?.InnerText ?? cell.InnerText;
            if (cell.DataType?.Value == CellValues.SharedString && sharedStrings != null
                && int.TryParse(value, out int idx))
            {
                return sharedStrings.ElementAt(idx).InnerText;
            }

            if (cell.DataType?.Value == CellValues.InlineString)
                return cell.InlineString?.InnerText ?? value ?? string.Empty;

            if (cell.DataType?.Value == CellValues.String)
                return value ?? string.Empty;

            if (cell.DataType?.Value == CellValues.Boolean)
                return value == "1" ? "TRUE" : "FALSE";

            if (IsDateLikeCell(cell, styles) &&
                double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var oaDate))
            {
                try
                {
                    var date = DateTime.FromOADate(oaDate);
                    return $"{date.Day}.{date.Month}.{date.Year % 100}";
                }
                catch
                {
                    // Si Excel tiene un serial inválido, devolvemos el valor crudo.
                }
            }

            return value ?? string.Empty;
        }

        private static bool IsDateLikeCell(Cell cell, Stylesheet? styles)
        {
            if (styles?.CellFormats == null || cell.StyleIndex == null)
                return false;

            var styleIndex = (int)cell.StyleIndex.Value;
            if (styleIndex < 0 || styleIndex >= styles.CellFormats.Count())
                return false;

            var cellFormat = styles.CellFormats.Elements<CellFormat>().ElementAt(styleIndex);
            var formatId = cellFormat.NumberFormatId?.Value ?? 0;

            if (IsBuiltInDateFormat(formatId))
                return true;

            if (styles.NumberingFormats == null)
                return false;

            var numberingFormat = styles.NumberingFormats.Elements<NumberingFormat>()
                .FirstOrDefault(n => n.NumberFormatId?.Value == formatId);
            var formatCode = numberingFormat?.FormatCode?.Value;
            return IsCustomDateFormat(formatCode);
        }

        private static bool IsBuiltInDateFormat(uint formatId)
        {
            return formatId is 14 or 15 or 16 or 17 or 18 or 19 or 20 or 21 or 22
                or 27 or 30 or 36 or 45 or 46 or 47 or 50 or 57;
        }

        private static bool IsCustomDateFormat(string? formatCode)
        {
            if (string.IsNullOrWhiteSpace(formatCode))
                return false;

            var code = formatCode.ToLowerInvariant();
            return (code.Contains('d') || code.Contains('m') || code.Contains('y'))
                && !code.Contains("[h]")
                && !code.Contains("[m]")
                && !code.Contains("[s]");
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
