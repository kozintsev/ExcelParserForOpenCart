using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;

namespace ExcelParserForOpenCart.Prices
{
    public class Rival : GeneralMethods
    {
        public Rival(object sender, DoWorkEventArgs e)
            : base(sender, e)
        {

        }
        /// <summary>
        /// Обработка ЕКБ_Прайс АвтоБРОНЯ_Игорь.xls
        /// </summary>
        /// <param name="row"></param>
        /// <param name="range"></param>
        public void AnalyzeBronya(int row, Range range)
        {
            Analyze(6, "15986394", 17, 4, row, range);
        }

        /// <summary>
        /// Обработка Подкрылки_Игорь.xls
        /// </summary>
        /// <param name="row"></param>
        /// <param name="range"></param>
        public void AnalyzePodkrilki(int row, Range range)
        {
            Analyze(4, "6697728", 6, 3, row, range);
        }
        /// <summary>
        /// Обработка Подлокотники_Игорь.xls
        /// </summary>
        /// <param name="row"></param>
        /// <param name="range"></param>
        public void AnalyzePodlokotniki(int row, Range range)
        {
            Analyze(4, "6697728", 9, 3, row, range);
        }
        /// <summary>
        /// Общий анализатор для всех прайсов Риваль
        /// </summary>
        /// <param name="startIndex"></param>
        /// <param name="codeColor"></param>
        /// <param name="columCost"></param>
        /// <param name="cat1nrow"></param>
        /// <param name="row"></param>
        /// <param name="range"></param>
        private void Analyze(int startIndex, string codeColor, int columCost, int cat1nrow, int row, Range range)
        {
            if (Worker.CancellationPending)
            {
                E.Cancel = true;
                ResultingPrice.Clear();
                return;
            }
            var countEmptyRow = 0;
            var icount = 0;
            var compareVendorCode = string.Empty;
            var unionDescription = string.Empty;
                ResultingPrice.Clear();
            var tmpResultingPrice = new List<OutputPriceLine>();
            var category1 = ConverterToString(range.Cells[cat1nrow, 1] as Range); //1 категория
            var category2 = string.Empty;
            // цикл для обработки прайс листа
            for (var i = startIndex; i < row; i++)
            {
                if (Worker.CancellationPending)
                {
                    E.Cancel = true;
                    ResultingPrice.Clear();
                    break;
                }
                var secRange = range.Cells[i, 1] as Range; //2 категория
                if (secRange != null)
                {
                    string str = ConverterToString(secRange.Value2);
                    var color = secRange.Interior.Color;
                    var sc = color.ToString();

                    if (sc == codeColor) // 2 категория
                    {
                        countEmptyRow = 0; //идет новая категория 2, зануляем счет на пустые строки
                        continue;
                    }
                    category2 = str;
                }
                if (secRange != null)
                {
                    string str = ConverterToString(secRange.Value2);
                    {
                        category2 = str;
                    }
                }
                var line = new OutputPriceLine
                {
                    Category1 = category1,
                    Category2 = category2
                };
                var vendorCode = ConverterToString(range.Cells[i, 5] as Range);
                if (compareVendorCode != vendorCode)
                {
                    compareVendorCode = vendorCode;
                    unionDescription = "<p>" + ConverterToString(range.Cells[i, 2] as Range).TrimEnd() + " (" + ConverterToString(range.Cells[i, 3] as Range) + ")" + "</p>";
                    line.ProductDescription = unionDescription;
                }
                else if (compareVendorCode == vendorCode)
                {
                    unionDescription += "<p>" + ConverterToString(range.Cells[i, 2] as Range).TrimEnd() + " (" + ConverterToString(range.Cells[i, 3] as Range) + ")" + "</p>";
                    if (vendorCode != "")
                    {
                        tmpResultingPrice[icount - 1].ProductDescription = unionDescription; //модифицируем
                    }
                    else if (string.IsNullOrEmpty(vendorCode)) countEmptyRow++;
                    if (countEmptyRow >= 3) break;  // выходить из цикла, после 3-й пустой строки
                    continue; // пропускаем строку
                }
                line.Name = ConverterToString(range.Cells[i, 4] as Range);
                var producer = GetManufacturers(line.Name);
                line.Producer = string.IsNullOrWhiteSpace(producer) ? GetManufacturers(line.Category2) : producer;
                line.Qt = "1000";
                if (string.IsNullOrEmpty(vendorCode) && !string.IsNullOrEmpty(line.Name))
                {
                    continue; // игнорировать строки без артикля
                }
                line.Cost = ConverterToString(range.Cells[i, columCost] as Range);
                line.VendorCode = vendorCode;
                if (string.IsNullOrEmpty(line.Name)) continue;
                tmpResultingPrice.Add(line); 
                icount++;
            }
            ResultingPrice.AddRange(Normalize(tmpResultingPrice).OrderBy(x => x.Category2).ToList()); //сортируем по категории 2
            tmpResultingPrice.Clear();
        }

        /// <summary>
        /// Выполняется "схлопывание" записей по одинаковому артикулу, категории2, цене, объединяются значения в ProductDescription
        /// </summary>
        /// <param name="tmpResultingPrice"></param>
        /// <returns></returns>
        private static IEnumerable<OutputPriceLine> Normalize(ICollection<OutputPriceLine> tmpResultingPrice)
        {
            var result = tmpResultingPrice.OrderBy(x => x.VendorCode).ThenBy(x => x.Category2).ToList();
            tmpResultingPrice.Clear();
            for (var i = 0; result.Count >= i; i++)
            {
                if (i == 0) { tmpResultingPrice.Add(result[i]); }
                if (i > 0 && result.Count != i)
                {
                    if (result[i].VendorCode == result[i - 1].VendorCode &&
                        result[i].Category2 == result[i - 1].Category2 && result[i].Cost == result[i - 1].Cost)
                    {
                        result[i - 1].ProductDescription += result[i].ProductDescription;//добавляем в предыдущую строку значение
                        result.RemoveAt(i);
                    }
                    else
                    {
                        tmpResultingPrice.Add(result[i]);
                    }
                }
            }
            result.Clear();
            return tmpResultingPrice;
        }
    }
}
