using System;
using ExcelParserForOpenCart.Prices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace ExcelParserForOpenCart
{
    public partial class ExcelParser : GeneralMethods
    {
        private static bool IsExcelInstall()
        {
            var hkcr = Registry.ClassesRoot;
            var excelKey = hkcr.OpenSubKey("Excel.Application");
            return excelKey != null;
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                if (obj != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            finally
            {
                GC.Collect();
            }
        }
        /// <summary>
        /// Определение типа прайс листа
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        private static EnumPrices DetermineTypeOfPriceList(Range range)
        {
            var str = ConverterToString(range.Cells[2, 3] as Range);
            if (str.Contains("Два Союза"))
                return EnumPrices.ДваСоюза;

            var str1 = ConverterToString(range.Cells[1, 1] as Range);
            var str2 = ConverterToString(range.Cells[1, 4] as Range);
            if (str1.Contains("Рисунок") && str2.Contains("Марка и модель автомобиля"))
                return EnumPrices.OJ;

            str1 = ConverterToString(range.Cells[9, 3] as Range);
            str2 = ConverterToString(range.Cells[11, 3] as Range);

            if (str1.Contains("Прайс-лист") && str2.Contains("Наименование товаров"))
                return EnumPrices.Autogur73;

            str1 = ConverterToString(range.Cells[2, 2] as Range);

            if (str1.Contains("Левандовская"))
                return EnumPrices.Левандовская;

            str1 = ConverterToString(range.Cells[9, 3] as Range);
            str2 = ConverterToString(range.Cells[9, 7] as Range);
            var range1 = range.Cells[9, 7] as Range;
            if (range1 == null) return EnumPrices.Неизвестный;
            var color = range1.Interior.Color;
            var sc = color.ToString();

            if (str1.Contains("Продукция") && str2.Contains("Безналичный расчет") && sc == "11842740")
                return EnumPrices.ПТГрупп;

            str1 = ConverterToString(range.Cells[3, 2] as Range);

            if (str1.Contains("ВЕНТУРИ"))
                return EnumPrices.Автовентури;


            return EnumPrices.Неизвестный;
        }
    }
}
