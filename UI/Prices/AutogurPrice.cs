using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    public class AutogurPrice : GeneralMethods
    {
        private readonly BackgroundWorker _worker;
        private readonly DoWorkEventArgs _e;

        public AutogurPrice()
        {
            
        }

        public AutogurPrice(object sender, DoWorkEventArgs e)
        {
            _worker = sender as BackgroundWorker;
            _e = e;
        }

        /// <summary>
        /// Прайс ИП Пьянов С.Г. Autogur73.ru
        /// </summary>
        /// <param name="row"></param>
        /// <param name="range"></param>
        public void Analyze(int row, Range range)
        {
            if (_worker.CancellationPending)
            {
                _e.Cancel = true;
                return;
            }
            var category1 = string.Empty;
            var category2 = string.Empty;
            var code = string.Empty;
            var vendorCode = string.Empty;
            var pair = false;
            List.Clear();
            // список имён с одинковым артикулем
            var list = new List<PairProductAndCost>();
            const string pattern = "(\\d+\\.\\s?)";
            for (var i = 13; i < row; i++)
            {
                if (_worker.CancellationPending)
                {
                    _e.Cancel = true;
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
                //var cost = ConverterToString(range.Cells[i, 5] as Range);
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

                if (string.IsNullOrEmpty(vendorCode) && string.IsNullOrEmpty(code) && string.IsNullOrEmpty(line.Name))
                    break; // выходить из цикла
                if (string.IsNullOrEmpty(vendorCode) && string.IsNullOrEmpty(code) && !string.IsNullOrEmpty(line.Name))
                    continue; // игнорировать строки без кода и артикля
                if (!string.IsNullOrEmpty(line.Name))
                    List.Add(line);
            }
        }
        /// <summary>
        /// Парсинг опции для прайса ИП Пьянов С.Г. Autogur73.ru
        /// </summary>
        /// <param name="list"></param>
        /// <param name="name"></param>
        /// <param name="options"></param>
        /// <param name="costs"></param>
        private static void GetNameAndOptionFromAutogur73(IReadOnlyList<PairProductAndCost> list,
            out string name, out string options, out string costs)
        {
            var i = 0;
            var maxStr = "";
            var minStr = "";
            var @case = 1;
            decimal cost = 0;
            costs = "";
            options = "";
            foreach (var s in list)
            {
                if (maxStr.Length < s.Name.Length)
                    maxStr = s.Name;
            }
            foreach (var s in list)
            {
                if (i == 0)
                {
                    minStr = s.Name;
                    i++;
                    continue;
                }
                if (s.Name.Length < minStr.Length)
                    minStr = s.Name;
            }
            if (maxStr.Length - minStr.Length < 5) @case = 3;
            name = minStr;
            if (@case == 1)
            {
                options = string.Empty;
                i = 0;
                var isFirstItem = true;
                foreach (var item in list)
                {
                    var option = item.Name.Replace(minStr, string.Empty).Replace(",", "").Trim();
                    if (isFirstItem)
                    {
                        isFirstItem = false;
                        cost = item.Cost;
                    }
                    if (option.Length > 19)
                    {
                        @case = 2;
                        break;
                    }
                    if (string.IsNullOrWhiteSpace(option)) continue;
                    if (i == 0)
                    {
                        options = option.Trim();
                        var diff = item.Cost - cost;
                        costs = diff.ToString(CultureInfo.CurrentCulture);
                    }
                    else
                    {
                        options += " ; " + option.Trim();
                        var diff = item.Cost - cost;
                        costs += " ; " + diff.ToString(CultureInfo.CurrentCulture);
                    }
                    i++;
                }
                options = options.Trim();
            }
            if (@case == 2)
            {
                options = string.Empty;
                var words = minStr.Split(new[] { ' ', ',', ':', '?', '!', ')' }, StringSplitOptions.RemoveEmptyEntries);
                i = 0;
                var isFirstItem = true;
                foreach (var item in list)
                {
                    if (item.Name == minStr)
                    {
                        if (isFirstItem)
                        {
                            isFirstItem = false;
                            cost = item.Cost;
                            costs = "0";
                        }
                        continue;
                    }
                    var option = item.Name.Replace(")", "");
                    foreach (var word in words)
                    {
                        if (word.Length == 1)
                            continue;
                        option = option.Replace(word, "");
                    }
                    option = option.Replace(",", "").Replace("(", "");
                    if (i == 0)
                    {
                        options = option.Trim();
                        var diff = item.Cost - cost;
                        costs = diff.ToString(CultureInfo.CurrentCulture);
                    }
                    else
                    {
                        options += " ; " + option.Trim();
                        var diff = item.Cost - cost;
                        costs += " ; " + diff.ToString(CultureInfo.CurrentCulture);
                    }
                    i++;
                }
                options = options.Trim();
            }
            if (@case != 3) return;
            i = 0;
            foreach (var item in list)
            {
                var j = item.Name.LastIndexOf(',') + 1;
                var option = "";
                for (var k = j; k < item.Name.Length; k++)
                {
                    option += item.Name[k];
                }
                if (i == 0)
                {
                    options = option.Trim();
                    cost = item.Cost;
                    costs = "0";
                    name = name.Replace(option, "").Replace(",", "").Trim();
                }
                else
                {
                    options += " ; " + option.Trim();
                    var diff = item.Cost - cost;
                    costs += " ; " + diff.ToString(CultureInfo.CurrentCulture);
                    name = name.Replace(option, "").Replace(",", "").Trim();
                }
                i++;
            }
        }
    }
}
