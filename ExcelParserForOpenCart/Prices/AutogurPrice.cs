using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    public class AutogurPrice : GeneralMethods
    {
        public AutogurPrice(object sender, DoWorkEventArgs e)
            : base(sender, e)
        {
           
        }

        /// <summary>
        /// Прайс ИП Пьянов С.Г. Autogur73.ru
        /// </summary>
        /// <param name="row"></param>
        /// <param name="range"></param>
        public void Analyze(int row, Range range)
        {
            if (Worker.CancellationPending)
            {
                E.Cancel = true;
                ResultingPrice.Clear();
                return;
            }
            var category1 = string.Empty;
            var category2 = string.Empty;
            var code = string.Empty;
            var vendorCode = string.Empty;
            var pair = false;
            ResultingPrice.Clear();
            // список имён с одинковым артикулем
            var list = new List<PairProductAndCost>();
            const string pattern = "(\\d+\\.\\s?)";
            for (var i = 13; i < row; i++)
            {
                if (Worker.CancellationPending)
                {
                    E.Cancel = true;
                    ResultingPrice.Clear();
                    break;
                }
                var line = new OutputPriceLine();
                var theRange = range.Cells[i, 3] as Range;
                if (theRange != null)
                {
                    string str = ConverterToString(theRange.Value2);
                    var color = theRange.Interior.Color;
                    var sc = color.ToString();
                    if (sc == "8765644") // 1 категория
                    {
                        category1 = Regex.Replace(str, pattern, string.Empty);
                        category2 = string.Empty;
                        continue;
                    }
                    if (sc == "9951719") // 2 категория
                    {
                        category2 = str;
                        continue;
                    }
                    if (sc == "12710911") continue; // пока предлагаю эту графу пропускать, это как бы 3 категория
                }
                line.Category1 = category1;
                line.Category2 = category2;
                code = ConverterToString(range.Cells[i, 1] as Range);
                vendorCode = ConverterToString(range.Cells[i, 2] as Range);
                // получить артикул следующей строки для сравнения
                var tempVendorCode = ConverterToString(range.Cells[i + 1, 2] as Range);
                var tempName = ConverterToString(range.Cells[i, 3] as Range);
                var tempName2 = ConverterToString(range.Cells[i + 1, 3] as Range);
                var cost1 = ConverterToDecimal(range.Cells[i, 5] as Range);
                var cost2 = ConverterToDecimal(range.Cells[i + 1, 5] as Range);
                if (!pair)
                {
                    list.Clear();
                    list.Add(new PairProductAndCost
                    {
                        Name = tempName,
                        Cost = cost1
                    });
                }
                if (string.IsNullOrWhiteSpace(vendorCode) && string.IsNullOrWhiteSpace(tempVendorCode) && string.IsNullOrWhiteSpace(tempName))
                    break;
                //todo: протестировать случай если артикуль не заполнен
                if (tempVendorCode == vendorCode && !string.IsNullOrWhiteSpace(vendorCode))
                {
                    // дублирование
                    list.Add(new PairProductAndCost
                    {
                        Name = tempName2,
                        Cost = cost2
                    });
                    pair = true;
                    continue;
                }
                // получить из списка опции и имя
                if (list.Count >= 2)
                {
                    string name, options, costs;
                    GetNameAndOptionFromAutogur73(list, out name, out options, out costs);
                    line.Name = name;
                    line.Option = options;
                    line.Cost = list[0].Cost.ToString(CultureInfo.CurrentCulture);
                    line.PlusThePrice = costs;
                }
                else
                {
                    line.Name = tempName;
                    line.Cost = cost1.ToString(CultureInfo.CurrentCulture);
                }
                list.Clear();
                pair = false;
                line.VendorCode = string.IsNullOrEmpty(vendorCode) ? code : vendorCode;
                line.Qt = "1000";
                line.Producer = GetManufacturers(line.Name);

                if (string.IsNullOrEmpty(vendorCode) && string.IsNullOrEmpty(code) && string.IsNullOrEmpty(line.Name))
                    break; 
                if (string.IsNullOrEmpty(vendorCode) && string.IsNullOrEmpty(code) && !string.IsNullOrEmpty(line.Name))
                    continue; // игнорировать строки без кода и артикля
                if (!string.IsNullOrEmpty(line.Name))
                    ResultingPrice.Add(line);
            }
        }
        /// <summary>
        /// Парсинг опции для прайса ИП Пьянов С.Г. Autogur73.ru
        /// </summary>
        /// <param name="list"></param>
        /// <param name="name"></param>
        /// <param name="options"></param>
        /// <param name="diffCosts"></param>
        private static void GetNameAndOptionFromAutogur73(IReadOnlyList<PairProductAndCost> list,
            out string name, out string options, out string diffCosts)
        {
            var isFirst = true;
            var maxStr = "";
            var minStr = "";
            var @case = 1;
            decimal cost = 0;
            var countOpt = 0;
            diffCosts = "";
            options = "";
            var separator = new[] {' ', ',', ';', ':', '?', '!', ')', '('};
            foreach (var s in list)
            {
                if (maxStr.Length < s.Name.Length)
                    maxStr = s.Name;
            }
            foreach (var s in list)
            {
                if (isFirst)
                {
                    minStr = s.Name;
                    isFirst = false;
                    continue;
                }
                if (s.Name.Length < minStr.Length)
                    minStr = s.Name;
            }
            if (maxStr.Length - minStr.Length < 5) @case = 2;
            name = minStr;
            var wordsMinStr = minStr.Split(separator, StringSplitOptions.RemoveEmptyEntries);
            if (@case == 1)
            {
                options = string.Empty;
                isFirst = true;
                foreach (var item in list.Where(item => item.Name == minStr))
                {
                    cost = item.Cost;
                    break;
                }
                foreach (var item in list)
                {
                    if (item.Name == minStr)
                    {
                        if (isFirst)
                        {
                            options = "Нет";
                            diffCosts = "0";
                            countOpt++;
                        }
                        else
                        {
                            options += ";Нет";
                            diffCosts = ";0";
                            countOpt++;
                        }
                        isFirst = false;
                        continue;
                    }
                    var tmpWords = item.Name.Split(separator, StringSplitOptions.RemoveEmptyEntries);
                    var option = tmpWords.Except(wordsMinStr).Aggregate("", (current, w) => current + (w + " ")).Trim();
                    if (option.Length > 19)
                    {
                        @case = 2;
                        break;
                    }
                    if (isFirst)
                    {
                        options = option.Trim();
                        diffCosts = GetDiffCost(item.Cost, cost);
                        isFirst = false;
                        countOpt++;
                    }
                    else
                    {
                        options += ";" + option.Trim();
                        diffCosts += ";" + GetDiffCost(item.Cost, cost);
                        countOpt++;
                    }
                }
                options = options.Trim();
            }
            if (@case == 2)
            {
                countOpt = 0;
                var totslStr = new List<string>();
                foreach (var item in list)
                {
                    if (item.Name == minStr) continue;
                    var tmpWords = item.Name.Split(separator, StringSplitOptions.RemoveEmptyEntries);
                    totslStr = wordsMinStr.Intersect(tmpWords).ToList();
                    break;
                }
                isFirst = true;
                foreach (var item in list)
                {
                    var tmpWords = item.Name.Split(separator, StringSplitOptions.RemoveEmptyEntries);
                    var option = tmpWords.Except(totslStr).Aggregate("", (current, w) => current + (w + " ")).Trim();
                    if (isFirst)
                    {
                        options = option.Trim();
                        cost = item.Cost;
                        diffCosts = "0";
                        isFirst = false;
                    }
                    else
                    {
                        options += ";" + option.Trim();
                        diffCosts += ";" + GetDiffCost(item.Cost, cost);
                    }
                    countOpt++;
                    if (!string.IsNullOrWhiteSpace(option)) name = name.Replace(option, "").Replace(",", "");
                }
            }
            if (countOpt == 2 && Global.IgnorOptions.Any(options.Contains))
            {
                options = "";
                diffCosts = "";
            }
        }

        private static string GetDiffCost(decimal itemCost, decimal cost)
        {
            var diff = itemCost - cost;
            var postfix = diff > 0 ? "+" : "-";
            if (diff == 0) postfix = "";
            diff = Math.Abs(diff);
            return diff.ToString(CultureInfo.CurrentCulture) + postfix;
        }
    }
}
