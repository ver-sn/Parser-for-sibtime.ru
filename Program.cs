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
            string[] articles = { "t1064071605100", "t1064071603100" };

            foreach (string article in articles)
            {
                var html = @"https://www.tissotwatches.com/ru-ru/shop/" + article + ".html";
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

                string myString = "Артикул";
                string gender = "Пол";
                string waterResist = "Водонепроницаемость";
                string material = "Материал корпуса";
                string width = "Ширина";
                string thickness = "Толщина";
                string glass = "Стекло";
                string dial = "Цвет циферблата";
                string mechanism = "Механизм";
                string belt = "Оформление ремешка/браслета";
                string colorBelt = "Цвет ремешка/браслета";

                var Result = new List<string>();

                foreach (var n in Nodes)
                {
                    if (n.Contains(myString)|| n.Contains(gender) || n.Contains(waterResist) || n.Contains(material) || n.Contains(width) || n.Contains(thickness) || n.Contains(glass) || n.Contains(dial) || n.Contains(mechanism) || n.Contains(belt) || n.Contains(colorBelt))
                    {
                        Result.Add(n);
                    }
                }

                foreach (var r in Result)
                {
                    Console.WriteLine(r);
                }
                Console.WriteLine(Result.Count);
                Console.ReadKey();


                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    //Opening an existing Excel file
                    FileInfo file = new FileInfo(@"D:\file\tissot.xlsx");

                    //create a WorkSheet
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                    //add all the content from the List collection
                    for (int x = 1; x < articles.Length+1;  x++)
                    {
                        worksheet.Cells[1, x].LoadFromCollection(Result);
                        excelPackage.SaveAs(file);
                    }
                    
                }

                Console.WriteLine("файл excel сохранен");

            }
                       
            //WebClient wc = new WebClient();
            //string[] articlesBase = { "T106.407.16.051.00", "T106.407.16.031.00" };
            //foreach (string articleBase in articlesBase)
            //{
            //    string path = @"https://www.tissotwatches.com/media/shop/catalog/product/T/0/" + articleBase + ".png";//создаем переменую с урл файла
            //    wc.DownloadFileAsync(new Uri(path), @"D:\Downloads\tissot\" + System.IO.Path.GetFileName(path));//скачиваем файл по указанному пути в указанное место на диске С
            //}
            
            //Console.WriteLine("Картинки сохранены");
            //Console.Read();
        }
    }
}
