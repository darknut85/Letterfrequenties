using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using GemBox.Document;

namespace Letterfrequenties
{
    class Program
    {
        static void Main(string[] args)
        {
            /////////////////////////////////////////////////////////////////////////
            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // Load Word document from file's path. (dit is de enige regel code die het programma nodig heeft. De rest is het lezen van text)
            // Deze package kan alleen casual gebruikt worden tot aan 20 alinea's. Daarna is hij niet langer gratis.
            var document = DocumentModel.Load("toRead.docx");

            // Get Word document's plain text.
            string text = document.Content.ToString();

            // Get Word document's count statistics.
            char write = 'x';
            int charactersCount = text.Replace(Environment.NewLine, string.Empty).Length;
            int wordsCount = document.Content.CountWords();
            int paragraphsCount = document.GetChildElements(true, ElementType.Paragraph).Count();
            int pageCount = document.GetPaginator().Pages.Count;
            int countLetters = text.Count(f => f == write);

            // Display file's count statistics.
            Console.WriteLine($"Letter count: {countLetters} for letter " + write);
            Console.WriteLine($"Characters count: {charactersCount}");
            Console.WriteLine($"     Words count: {wordsCount}");
            Console.WriteLine($"Paragraphs count: {paragraphsCount}");
            Console.WriteLine($"     Pages count: {pageCount}");
            Console.WriteLine();

            // Display file's text content.
            Console.WriteLine(text);
        }
    }
}
