using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace WordToPdfConverter
{
    class Program
    {
        static void Main()
        {
            Console.Write("Digite o caminho do arquivo Word (.docx): ");
            string wordFilePath = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(wordFilePath) || !File.Exists(wordFilePath))
            {
                Console.WriteLine("Arquivo inválido ou não encontrado.");
                return;
            }

            string pdfFilePath = Path.ChangeExtension(wordFilePath, ".pdf");

            try
            {
                ConvertWordToPdf(wordFilePath, pdfFilePath);
                Console.WriteLine($"Conversão concluída! Arquivo salvo em: {pdfFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro na conversão: {ex.Message}");
            }
        }

        static void ConvertWordToPdf(string wordPath, string pdfPath)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = null;

            try
            {
                wordDoc = wordApp.Documents.Open(wordPath);
                wordDoc.ExportAsFixedFormat(pdfPath, Word.WdExportFormat.wdExportFormatPDF);
            }
            finally
            {
                wordDoc?.Close(false);
                wordApp.Quit();
            }
        }
    }
}
