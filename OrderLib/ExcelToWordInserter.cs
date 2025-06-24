using System;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OrderLib
{
    public class ExcelToWordInserter
    {
        public void Insert(string excelPath, string wordPath)
        {
            if (!File.Exists(excelPath)) throw new FileNotFoundException("Excel file not found", excelPath);
            if (!File.Exists(wordPath)) throw new FileNotFoundException("Word file not found", wordPath);

            try
            {
                using var workbook = new XLWorkbook(excelPath);
                var worksheet = workbook.Worksheets.Worksheet(1);
                if (worksheet == null) throw new InvalidOperationException("No worksheet found in Excel file");

                using var doc = WordprocessingDocument.Open(wordPath, true);
                var body = doc.MainDocumentPart?.Document?.Body;
                if (body == null) throw new InvalidOperationException("Invalid Word document structure");

                var table = new Table();

                foreach (var row in worksheet.RangeUsed().Rows())
                {
                    var wordRow = new TableRow();
                    foreach (var cell in row.Cells())
                    {
                        if (cell == null || cell.GetString() == null) continue;
                        var wordCell = new TableCell(new Paragraph(new Run(new Text(cell.GetString()))));
                        wordRow.Append(wordCell);
                    }
                    if (wordRow.ChildElements.Count > 0) table.Append(wordRow);
                }

                body.Append(table);
                doc.MainDocumentPart.Document.Save();
            }
            catch (IOException ex)
            {
                throw new IOException("Error inserting Excel data into Word: " + ex.Message, ex);
            }
        }
    }
}