using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OrderLib
{
    public class OrderProcessor
    {
        private readonly TxtReader _xmlParser;
        private readonly ExcelToWordInserter _excelInserter;
        private readonly WordTexter _writer;

        public OrderProcessor(TxtReader xmlParser, ExcelToWordInserter excelInserter, WordTexter writer)
        {
            _xmlParser = xmlParser ?? throw new ArgumentNullException(nameof(xmlParser));
            _excelInserter = excelInserter ?? throw new ArgumentNullException(nameof(excelInserter));
            _writer = writer ?? throw new ArgumentNullException(nameof(writer));
        }

        public OrderProcessor() : this(new TxtReader(), new ExcelToWordInserter(), new WordTexter())
        {
        }

        public void Process(string inputExcel, string outputWord, string inputXml)
        {
            if (!File.Exists(inputExcel)) throw new FileNotFoundException("Excel file not found", inputExcel);
            if (!File.Exists(inputXml)) throw new FileNotFoundException("XML file not found", inputXml);

            try
            {
                string textContent = _xmlParser.GetText(inputXml) ?? string.Empty;
                
                _writer.CreateWordWithText(outputWord, textContent);

                _excelInserter.Insert(inputExcel, outputWord);
            }
            catch (IOException ex)
            {
                throw new IOException("Error processing files: " + ex.Message, ex);
            }
            catch (Exception ex)
            {
                throw new Exception("Unexpected error during processing: " + ex.Message, ex);
            }
        }
    }
    public class TxtReader
    {
        public string GetText(string inputXml) => File.ReadAllText(inputXml); 
    }

    public class WordTexter
    {
        public void CreateWordWithText(string outputWord, string textContent)
        {
            // Ensure output path ends with .docx
            if (!outputWord.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                outputWord += ".docx";
            }

            try
            {
                // Create or overwrite the Word document
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputWord, WordprocessingDocumentType.Document))
                {
                    // Add main document part
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    // Add paragraph with text
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text(textContent));
                    wordDoc.Dispose();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error creating Word document: " + ex.Message, ex);
            }
        }
    }
}