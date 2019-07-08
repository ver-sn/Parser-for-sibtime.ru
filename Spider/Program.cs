using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using HtmlAgilityPack;
using OfficeOpenXml;
using System.IO;
using System.Net;




namespace Spider
{
    class Program
    {
        static void Main(string[] args)
        {
            //var art = "t1064071605100.html";
            string[] arts = { "t1064071605100", "t1166171603700" };

            foreach (string art in arts)
            {
                var html = @"https://www.tissotwatches.com/ru-ru/shop/" + art + ".html";

                HtmlWeb web = new HtmlWeb();

                var htmlDoc = web.Load(html);

                var htmlNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='specs-table-1024']");
                var Nodes = new List<string>();//создали лист


                foreach (var node in htmlNodes)
                {
                    var res = node.InnerHtml.ToString().Replace("<table class=\"specs-table\">", "").Replace("\t", "").Replace("\n", "").Replace("\r", "")
                        .Replace("<tr>", "").Replace("</tr>", "#").Split('#');

                    foreach (var i in res)
                    {
                        var t = i.Replace("<td>", "").Replace("</td>", "#").Split('#');

                        try
                        {
                            var entity = t[0] + " : " + t[1];
                            Nodes.Add(entity);
                        }
                        catch (Exception)
                        {
                            var s = "Wrong Entity : " + t;
                            Nodes.Add(s);
                        }

                    }


                }

                //foreach (var node in Nodes)
                //{
                //    Console.WriteLine(node);
                //}

                //Console.ReadKey();


                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    //create a WorkSheet
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");


                    //add all the content from the List collection, starting at cell A1
                    worksheet.Cells[1, 1].LoadFromCollection(Nodes);

                    FileInfo fi = new FileInfo(@"D:\file\" + art + ".xlsx");
                    excelPackage.SaveAs(fi);
                }

                Console.WriteLine("файл excel сохранен");

            }

            WebClient wc = new WebClient();
            string path = "https://www.tissotwatches.com/media/shop/catalog/product/T/0/T099.207.22.118.01.png";//создаем переменую с урл файла
            wc.DownloadFileAsync(new Uri(path), @"D:\Downloads\tissot\" + System.IO.Path.GetFileName(path));//скачиваем файл по указанному пути в указанное место на диске С

            Console.WriteLine("Картинка сохранена");
            Console.Read();
        }
    }
}
