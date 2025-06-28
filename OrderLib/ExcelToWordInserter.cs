using ClosedXML.Excel;
using DocumentFormat.OpenXml;
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
            using var workbook = new XLWorkbook(excelPath);
            var worksheet = workbook.Worksheets.Worksheet(1);

            using var doc = WordprocessingDocument.Open(wordPath, true);
            var body = doc.MainDocumentPart.Document.Body;

            var table = new Table();

            // Apply a table style (default Word table style)
            TableProperties tblProps = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 }
                )
            );
            table.AppendChild(tblProps);

            bool isFirstRow = true;
            foreach (var row in worksheet.RangeUsed().Rows())
            {
                var wordRow = new TableRow();
                foreach (var cell in row.Cells())
                {
                    Run run = new Run(new Text(cell.GetValue<string>()));

                    if (isFirstRow)
                    {
                        run.RunProperties = new RunProperties(new Bold());
                    }

                    var wordCell = new TableCell(new Paragraph(run));
                    wordRow.Append(wordCell);
                }
                table.Append(wordRow);
                isFirstRow = false;
            }

            body.Append(table);
            doc.MainDocumentPart.Document.Save();
        }
    }
}