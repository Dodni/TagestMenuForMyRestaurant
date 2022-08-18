using System.IO;
using SautinSoft.Document;
using System.Linq;
using System.Text.RegularExpressions;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            FindAndReplace();
        }

        /// <summary>
        /// Find and replace a text using ContentRange.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/find-replace-content-net-csharp-vb.php
        /// </remarks>
        public static void FindAndReplace()
        {
            // Path to a loadable document.
            string loadPath = @"template.docx";
            //string loadPath = @"..\..\critique.docx";

            // Load a document intoDocumentCore.
            DocumentCore dc = DocumentCore.Load(loadPath);

            //5 words to replace
            Regex appetizer = new Regex(@"<appetizer>", RegexOptions.IgnoreCase);
            Regex suppe = new Regex(@"<suppe>", RegexOptions.IgnoreCase);
            Regex salad = new Regex(@"<salad>", RegexOptions.IgnoreCase);
            Regex menua = new Regex(@"<menua>", RegexOptions.IgnoreCase);
            Regex menub = new Regex(@"<menub>", RegexOptions.IgnoreCase);
            Regex dessert = new Regex(@"<dessert>", RegexOptions.IgnoreCase);
            Regex day = new Regex(@"<day>", RegexOptions.IgnoreCase);
            Regex date = new Regex(@"<date>", RegexOptions.IgnoreCase);


            string appetizerStr;
            string suppeStr = "";
            string saladStr = "";
            string menuaStr = "";
            string menubStr = "";
            string dessertStr = "";
            string dayStr = "";
            string dateStr = "";


            Console.WriteLine("Geben Sie den Namen der Vorspeise ein:");
            appetizerStr = Console.ReadLine();
            Console.WriteLine("Das hat er geschrieben: " + appetizerStr + "\n");

            Console.WriteLine("Geben Sie den Namen der Suppe ein:");
            suppeStr = Console.ReadLine();
            Console.WriteLine("Das hat er geschrieben: " + suppeStr + "\n");

            Console.WriteLine("Geben Sie den Namen des Salats ein:");
            saladStr = Console.ReadLine();
            Console.WriteLine("Das hat er geschrieben: " + saladStr + "\n");

            Console.WriteLine("Geben Sie einen Namen für Menü A ein:");
            menuaStr = Console.ReadLine();
            Console.WriteLine("Das hat er geschrieben: " + menuaStr + "\n");

            Console.WriteLine("Geben Sie einen Namen für Menü B ein.:");
            menubStr = Console.ReadLine();
            Console.WriteLine("Das hat er geschrieben: " + menubStr + "\n");

            Console.WriteLine("Geben Sie den Namen des Desserts ein:");
            dessertStr = Console.ReadLine();
            Console.WriteLine("Das hat er geschrieben: " + menubStr + "\n");

            Console.WriteLine("Geben Sie den Namen des Tages ein:");
            dayStr = Console.ReadLine();
            Console.WriteLine("Das hat er geschrieben: " + dayStr + "\n");

            Console.WriteLine("Geben Sie das Datum ein:");
            dateStr = Console.ReadLine();
            Console.WriteLine("Das hat er geschrieben: " + dateStr + "\n");


            // Please note, Reverse() makes sure that action replace not affects to Find().
            foreach (ContentRange item in dc.Content.Find(appetizer).Reverse())
            {
                item.Replace(appetizerStr);
            }

            foreach (ContentRange item in dc.Content.Find(suppe).Reverse())
            {
                item.Replace(suppeStr);
            }

            foreach (ContentRange item in dc.Content.Find(salad).Reverse())
            {
                item.Replace(saladStr);
            }

            foreach (ContentRange item in dc.Content.Find(menua).Reverse())
            {
                item.Replace(menuaStr);
            }

            foreach (ContentRange item in dc.Content.Find(menub).Reverse())
            {
                item.Replace(menubStr);
            }

            foreach (ContentRange item in dc.Content.Find(dessert).Reverse())
            {
                item.Replace(dessertStr);
            }

            foreach (ContentRange item in dc.Content.Find(day).Reverse())
            {
                item.Replace(dayStr);
            }

            foreach (ContentRange item in dc.Content.Find(date).Reverse())
            {
                item.Replace(dateStr);
            }

            // Save our document into PDF format.
            string savePath = "Replaced.pdf";
            dc.Save(savePath, new PdfSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(loadPath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(savePath) { UseShellExecute = true });
        }
    }
}