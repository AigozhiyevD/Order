using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main(string[] args)
    {
        string inputExcel = GetArg(args, "--input-excel");
        string outputWord = GetArg(args, "--output-word");
        string xmlTextPath = GetArg(args, "--text");

        Console.WriteLine("Reading XML text...");
        string textContent = ReadTextFromXml(xmlTextPath);

        Console.WriteLine("Reading Excel...");
        DataTable table = ReadExcelAsDataTable(inputExcel);

        Console.WriteLine("Writing Word...");
        WriteToWord(outputWord, textContent, table);

        Console.WriteLine("Done.");
    }

    static string GetArg(string[] args, string name)
    {
        for (int i = 0; i < args.Length - 1; i++)
        {
            if (args[i] == name)
                return args[i + 1];
        }
        throw new Exception("Missing required argument: " + name);
    }

    static string ReadTextFromXml(string path)
    {
        XmlDocument doc = new XmlDocument();
        doc.Load(path);
        XmlNode node = doc.SelectSingleNode("//text");
        return node != null ? node.InnerText : "";
    }

    static DataTable ReadExcelAsDataTable(string filePath)
    {
        string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"";
        using (OleDbConnection conn = new OleDbConnection(connStr))
        {
            conn.Open();
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString();

            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "]", conn);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }
    }

    static void WriteToWord(string filePath, string text, DataTable table)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            // Add XML text paragraph
            Paragraph para = new Paragraph(new Run(new Text(text)));
            body.Append(para);

            // Create table
            Table wordTable = new Table();
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
            wordTable.AppendChild(tblProps);
            bool isFirstRow = true;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                TableRow tr = new TableRow();
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    string val = table.Rows[i][j].ToString();

                    Run run = new Run(new Text(val));
                    if (isFirstRow)
                    {
                        run.RunProperties = new RunProperties(new Bold());
                    }

                    var wordCell = new TableCell(new Paragraph(run));
                    tr.Append(wordCell);
                }
                wordTable.Append(tr);
                isFirstRow = false;
            }

            body.Append(wordTable);
        }
    }
}
