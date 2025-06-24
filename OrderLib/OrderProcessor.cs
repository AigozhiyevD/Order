using System;
using System.IO;
using OrderLib;

namespace OrderLib
{
    public class OrderProcessor
    {
        private readonly XmlParser _xmlParser;
        private readonly ExcelToWordInserter _excelInserter;
        private readonly WordWriter _writer;

        public OrderProcessor(XmlParser xmlParser, ExcelToWordInserter excelInserter, WordWriter writer)
        {
            _xmlParser = xmlParser ?? throw new ArgumentNullException(nameof(xmlParser));
            _excelInserter = excelInserter ?? throw new ArgumentNullException(nameof(excelInserter));
            _writer = writer ?? throw new ArgumentNullException(nameof(writer));
        }

        public OrderProcessor() : this(new XmlParser(), new ExcelToWordInserter(), new WordWriter())
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
    public class XmlParser
    {
        public string GetText(string inputXml) => File.ReadAllText(inputXml); 
    }

    public class WordWriter
    {
        public void CreateWordWithText(string outputWord, string textContent)
        {
            File.WriteAllText(outputWord, textContent);
        }
    }
}