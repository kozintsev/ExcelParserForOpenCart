using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    public class Autoventuri : GeneralMethods
    {
        private readonly List<Product> _list;
        public int CountOfLink;

        public Autoventuri(object sender, DoWorkEventArgs e) 
            : base(sender, e)
        {
            _list = new List<Product>();
            CountOfLink = 0;
        }

        public void ParseImg()
        {
            const string urlHost = "http://www.autoventuri.ru";
            _list.Clear();
            //получаем html страницу со всем барахлом включая результаты нашего поиска
            var doc = new HtmlWeb().Load(urlHost);
            var catalogs = doc.DocumentNode.SelectNodes("//*[@id=\"market\"]/div/div[2]/div[1]/div/div[2]/div/*/ul/*/a");
            foreach (var catalog in catalogs)
            {
                var uri = catalog.GetAttributeValue("href", "");
                GetListProducts(urlHost + uri);
            }
        }
        /// <summary>
        /// Обработка прайс-листа ВЕНТУРИ (ПРАЙС автовентури.xls)
        /// </summary>
        /// <param name="row"></param>
        /// <param name="range"></param>
        public void Analyze(int row, Range range)
        {
            if(Worker.CancellationPending)
            {
                E.Cancel = true;
                ResultingPrice.Clear();
                return;
            }
            var category1 = string.Empty;
            var category2 = string.Empty;
            ResultingPrice.Clear();
            // цикл для обработки прайс листа
            for (var i = 11; i < row; i++)
            {
                if (Worker.CancellationPending)
                {
                    E.Cancel = true;
                    ResultingPrice.Clear();
                    break;
                }

                var theRange = range.Cells[i, 2] as Range; //берем из 2 столбца
                if (theRange != null)
                {
                    string str = ConverterToString(theRange.Value2);
                    var color = theRange.Interior.Color;
                    var sc = color.ToString();
                    if (sc == "11842740") // 1 категория
                    {
                        category1 = str.TrimStart(' ');
                        category2 = string.Empty;
                        continue;
                    }
                    if (sc == "12829635") // 2 категория
                    {
                        category2 = str.TrimStart(' ');
                        continue;
                    }
                }

                var line = new OutputPriceLine
                {
                    Category1 = category1,
                    Category2 = category2
                };
                var vendorCode = ConverterToString(range.Cells[i, 3] as Range);

                line.Name = ConverterToString(range.Cells[i, 2] as Range).TrimStart(' '); // тримим пробелы вначале строки
                line.Producer = GetManufacturers(line.Name);

                if (string.IsNullOrEmpty(vendorCode) && !string.IsNullOrEmpty(line.Name))
                    continue; // игнорировать строки без артикля

                line.Cost = ConverterToString(range.Cells[i, 6] as Range);
                line.VendorCode = vendorCode;
                line.Qt = "1000";
                var foto = _list.Where(x => x.Num == vendorCode).Select(x => x.ImgUrl).FirstOrDefault();
                if (foto != null)
                {
                    line.Foto = foto;
                    CountOfLink++;
                }

                if (string.IsNullOrEmpty(vendorCode) && string.IsNullOrEmpty(line.Name))
                    break; // выходить из цикла

                if (!string.IsNullOrEmpty(line.Name))
                    ResultingPrice.Add(line);

            }
        }

        private void GetImages(HtmlDocument doc, string hostName)
        {
            //получаем список всех постов по нашему поиску, все остальное барахло мимо
            var posters =
                doc.DocumentNode.SelectNodes("//*[@id=\"wrap\"]/div/section/div[2]/div[6]/*");
            //получаемссылку на первый пост из нашего списка постов
            var i = 1;
            foreach (var poster in posters)
            {
                var num =
                    poster.SelectSingleNode("//*[@id=\"wrap\"]/div/section/div[2]/div[6]/div[" + i + "]/div/div[1]").InnerText;
                var urlImg = poster.SelectSingleNode("//*[@id=\"wrap\"]/div/section/div[2]/div[6]/div[" + i + "]/div/div[3]/a/img")
                    .GetAttributeValue("src", string.Empty);
                num = num.Replace("Арт.", "").Trim();
                var filename = Path.GetFileName(urlImg);
                if (filename != null)
                {
                    var s = filename[0].ToString() + filename[1] + filename[2];
                    // картинка в максимальном расширении
                    var imgUrl = string.Format("{0}/upload/iblock/{1}/{2}", hostName, s, filename);
                    _list.Add(new Product
                    {
                        Num = num,
                        ImgUrl = imgUrl
                    });
                }
                i++;
            }
        }

        private void GetListProducts(string url)
        {
            var myuri = new Uri(url);
            var pathQuery = myuri.PathAndQuery;
            var hostName = myuri.ToString().Replace(pathQuery, "");

            var doc = new HtmlWeb().Load(url.Trim());
            //проверять многостраничность
            var showall = doc.DocumentNode.SelectSingleNode("//*[@id=\"wrap\"]/div/section/div[2]/div[7]/noindex/a");
            if (showall != null)
            {
                var urlall = showall.GetAttributeValue("href", "");
                var addres = hostName + urlall;
                addres = addres.Replace("&amp;", "&");
                // HtmlAgilityPack не может с парамметрами
                var doc2 = new HtmlWeb().Load(addres, "GET");
                GetImages(doc2, hostName);
                return;
            }
            GetImages(doc, hostName);
        }
    }
}
