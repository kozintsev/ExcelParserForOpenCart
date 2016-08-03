﻿using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    public class For2Union : GeneralMethods
    {
        private readonly BackgroundWorker _worker;
        private readonly DoWorkEventArgs _e;

        public For2Union(object sender, DoWorkEventArgs e)
        {
            _worker = sender as BackgroundWorker;
            _e = e;
        }
        /// <summary>
        /// Обработка для прайса 2 союза
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
            List.Clear();
            for (var i = 9; i < row; i++)
            {
                if (_worker.CancellationPending)
                {                
                    _e.Cancel = true;
                    break;
                }
                var line = new OutputPriceLine();
                var str = string.Empty;
                var theRange = range.Cells[i, 1] as Range;
                if (theRange != null)
                {
                    str = ConverterToString(theRange.Value2);
                    var color = theRange.Interior.Color;
                    var sc = color.ToString();
                    if (sc == "0") // чёрный
                    {
                        category1 = str;
                        category2 = string.Empty;
                        continue;
                    }
                    if (sc == "8421504")
                    {
                        category2 = str;
                        continue;
                    }
                }
                line.Category1 = category1;
                line.Category2 = category2;
                line.VendorCode = ConverterToString(range.Cells[i, 3] as Range);
                line.Name = ConverterToString(range.Cells[i, 4] as Range);
                line.Qt = ConverterToString(range.Cells[i, 5] as Range);
                // todo: цена в USD может стоит её как-то обработать?
                line.Cost = ConverterToString(range.Cells[i, 6] as Range);

                if (!string.IsNullOrEmpty(line.VendorCode))
                    List.Add(line);
                if (string.IsNullOrEmpty(str)) break;
            }
        }
    }
}
