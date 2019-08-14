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




namespace Parser
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] vendorCodes = {
                "T006.207.11.036.00",
                "T006.207.11.038.00",
                "T006.207.11.058.00",
                "T006.207.11.116.00",
                "T006.207.16.038.00",
                "T006.207.22.038.00",
                "T006.207.22.116.00",
                "T006.414.36.443.00",
                "T006.428.11.038.02",
                "T006.428.22.038.02",
                "T006.428.36.058.02",
                "T019.430.11.041.00",
                "T035.207.11.031.00",
                "T035.207.22.031.00",
                "T035.207.36.061.00",
                "T035.210.11.031.00",
                "T035.210.16.031.01",
                "T035.614.11.051.01"};
            
            List<List<string>> Nodes = new List<List<string>>();
            List<string> Images = new List<string>();
                                 
            foreach (string vendorCode in vendorCodes)
            {
                var search = @"https://www.tissotwatches.com/ru-ru/shop/catalogsearch/result/?q=" + vendorCode;
                HtmlWeb web = new HtmlWeb();
                var searchDoc = web.Load(search);

                try
                {
                    var searchNode = searchDoc.DocumentNode.SelectSingleNode("//div[@class='hover-zone']/a").Attributes["href"].Value;
                           
                if (searchNode != null)
                {
                    Console.WriteLine(searchNode);
                }
                
                if (searchNode != null)
                {
                    var htmlDoc = web.Load(searchNode);
                    var htmlNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='specs-table-1024']");

                    var imageNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class='item']/a").Attributes["href"].Value;
                    Images.Add(imageNode);


                    var Product = new List<string>();//создали лист

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
                                Product.Add(entity);
                            }
                            catch (Exception)
                            {
                                var s = "Wrong Entity : " + t;
                                Product.Add(s);
                            }

                        }

                    }

                    
                    var Result = new List<string>();

                    foreach (var n in Product)
                    {
                        if (n.Contains(myString) || n.Contains(gender) || n.Contains(waterResist) || n.Contains(material) || n.Contains(width) || n.Contains(thickness) || n.Contains(glass) || n.Contains(dial) || n.Contains(mechanism) || n.Contains(belt) || n.Contains(colorBelt))
                        {
                            Result.Add(n);
                        }
                    }

                    Nodes.Add(Result);

                    //foreach (var r in Result)
                    //{
                    //    Console.WriteLine(r);
                    //}
                    //Console.WriteLine(Result.Count);
                    //Console.ReadKey();
                }
                }

                catch (Exception er)
                {
                    var a = er;
                }
            }

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                var iter = 1;
                foreach (var i in Nodes)
                {
                    var val = i.ToList();
                    //добавляем контент
                    worksheet.Cells[1, iter].LoadFromCollection(val);
                    iter++;
                }

                FileInfo fi = new FileInfo(@"D:\result.xlsx");
                excelPackage.SaveAs(fi);
            }

            Console.WriteLine("файл excel сохранен");

            WebClient wc = new WebClient();
            
            foreach (string img in Images)
            {
                string path = @img;//создаем переменую с урл файла 
                try
                {
                    wc.DownloadFileTaskAsync(new Uri(path), @"D:\tissot\" + System.IO.Path.GetFileName(path)).GetAwaiter().GetResult();//скачиваем файл по указанному пути в указанное место на диске С 
                }
                catch (Exception ex)
                {
                    var t = ex;
                }

            }

            Console.WriteLine("Картинка сохранена");
            Console.ReadKey();
        }
    }
}
