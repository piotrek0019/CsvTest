using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using System.IO;
using CsvHelper.Configuration.Attributes;
using CsvHelper.Configuration;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Diagnostics;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;

namespace CsvTest
{

    public class Foo
    {
        [Index(0)]
        public int Id { get; set; }
        [Index(1)]
        public string OrderNr { get; set; }
        [Index(2)]
        public int PalletNr { get; set; }
        [Index(3)]
        public int SiteNr { get; set; }
        [Index(4)]
        public string Item { get; set; }
        [Index(5)]
        public int Quantity { get; set; }
        [Index(6)]
        public string Type { get; set; }
        [Index(7)]
        public int Type2 { get; set; }
        [Index(8)]
        public string Type3 { get; set; }



    }
    class Program
    {
        static void Main(string[] args)
        {
           

            string items = "\n\nID: Items  -- Quantity";
            int suma = 0;
            int counter = 0;
            string orderAdress = "";
            Console.WriteLine("Type order number");
            string orderNr = Console.ReadLine();
           //orderAdress = "C:\\Portech\\PBS\\Pallets\\" + orderNr + ".csv";
            orderAdress = "C:\\Users\\piotr\\Desktop\\CSV\\" + orderNr + ".csv";

            if (string.IsNullOrEmpty(orderNr))
            {
                Console.WriteLine("Type order adress");
                orderAdress = Console.ReadLine();
            }

            Console.WriteLine("File's address: " + orderAdress);
            
            using (var reader = new StreamReader(orderAdress))
            {
                using (var csv = new CsvReader(reader))
                {
                    csv.Configuration.HasHeaderRecord = false;
                    //csv.Configuration.RegisterClassMap<FooMap>();
                    var record = new Foo();
                    var records = csv.EnumerateRecords(record);


                    var query =
                        (from pp in records
                         
                             select new
                             {
                                 Items = pp.Item,
                                 Quantity = pp.Quantity
                             }).GroupBy(c => c.Items).Select(c => new { Items = c.Key, Quantity = c.Sum(d => d.Quantity) }).ToList();

                    var query2 =
                        from xx in query
                        orderby xx.Items
                        select xx;

                    foreach (var x in query2)
                    {
                        counter += 1;
                        Console.WriteLine(counter + " " + x.Items + " " + x.Quantity);
                        items += "\n\n " + counter + ": " + x.Items + "------" + x.Quantity;
                        suma += x.Quantity;
                    }
                    items += "\n\n\n" + suma;

                    
                  
                    
                }
            }


            Document document = new Document();
            Section section = document.AddSection();
            section.AddParagraph("Wolf System");
            section.AddParagraph();
            Paragraph paragraph = section.AddParagraph();
            //paragraph.Format.Font.Color = Color.FromCmyk(100, 30, 20, 50);
            paragraph.AddFormattedText(orderNr, TextFormat.Underline);
            FormattedText ft = paragraph.AddFormattedText(items, TextFormat.Bold);
            ft.Font.Size = 14;

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(false);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            string filname = "WolfOrder.pdf";
            pdfRenderer.PdfDocument.Save(filname);
            Process.Start(filname);

            Console.ReadLine();
        }
    }
}
   






            
