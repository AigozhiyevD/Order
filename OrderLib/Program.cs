using System;
using OrderLib;

class Program
{
    static void Main()
    {
        var processor = new OrderProcessor();

        string pathToExcel = @"C:\Путь\к\файлу.xlsx";
        string pathToXml = @"C:\Путь\к\файлу.xml";
        string outputWord = @"C:\Путь\куда\сохранить.docx";

        try
        {
            processor.Process(pathToExcel, outputWord, pathToXml);
            Console.WriteLine("Обработка завершена успешно!");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Произошла ошибка: " + ex.Message);
        }
    }
}