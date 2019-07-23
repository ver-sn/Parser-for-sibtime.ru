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

            List<List<string>> Nodes = new List<List<string>>();

            foreach (string art in arts)
            {
                var html = @"https://www.tissotwatches.com/ru-ru/shop/" + art + ".html";

                HtmlWeb web = new HtmlWeb();

                var htmlDoc = web.Load(html);

                var htmlNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='specs-table-1024']");
                var Articul = new List<string>();//создали лист


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
                            Articul.Add(entity);
                        }
                        catch (Exception)
                        {
                            var s = "Wrong Entity : " + t;
                            Articul.Add(s);
                        }
                    }
                }



                Nodes.Add(Articul);



                //foreach (var node in Nodes)
                //{
                //    Console.WriteLine(node);
                //}

                //Console.ReadKey();


               

                //Console.WriteLine("файл excel сохранен");

            }

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //create a WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                var iter = 1;
                foreach (var i in Nodes)
                {
                    var val = i.ToList();
                    //add all the content from the List collection, starting at cell A1
                    worksheet.Cells[1, iter].LoadFromCollection(val);
                    iter++;
                }

                FileInfo fi = new FileInfo(@"D:\file\result.xlsx");
                excelPackage.SaveAs(fi);
            }

            //WebClient wc = new WebClient();
            //string path = "https://www.tissotwatches.com/media/shop/catalog/product/T/0/T099.207.22.118.01.png";//создаем переменую с урл файла
            //wc.DownloadFileAsync(new Uri(path), @"D:\Downloads\tissot\" + System.IO.Path.GetFileName(path));//скачиваем файл по указанному пути в указанное место на диске С


            WebClient wc = new WebClient();
            string[] articlesBase = { "T106.407.16.051.00", "T106.407.16.031.00" };
            foreach (string articleBase in articlesBase)
            {
                string path = @"https://www.tissotwatches.com/media/shop/catalog/product/T/1/" + articleBase + ".png";//создаем переменую с урл файла 
                try
                {
                   wc.DownloadFileTaskAsync(new Uri(path), @"D:\Downloads\tissot\" + System.IO.Path.GetFileName(path)).GetAwaiter().GetResult();//скачиваем файл по указанному пути в указанное место на диске С 
                }
                catch (Exception ex)
                {
                    var t = ex;
                }
               
            }

            Console.WriteLine("Картинка сохранена");
            Console.Read();
        }
    }
}
