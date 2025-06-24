using System;
using OrderLib;

class Program
{
    static void Main(string[] args)
    {
        try
        {
            var processor = new OrderProcessor();
            processor.Process("input.xlsx", "output.docx", "input.xml");
            Console.WriteLine("Processing completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}