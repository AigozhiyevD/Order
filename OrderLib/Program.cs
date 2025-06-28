using System;
using OrderLib;

class Program
{
    static void Main(string [] args)
    {
        foreach (var arg in args)
        {
            Console.WriteLine();
        }
        if (args.Length != 3)
        {
            Console.WriteLine("Not enough args provided");
            return;
        }
        
        var processor = new OrderProcessor();

        string pathToExcel = args[2];
        string pathToXml = args[1];
        string outputWord = args[0];
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