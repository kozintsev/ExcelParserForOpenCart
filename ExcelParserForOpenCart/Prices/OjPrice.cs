using System;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    public class OjPrice : GeneralMethods
    {
        public event Action<string> OnMsg;

        public OjPrice(object sender, DoWorkEventArgs e)
            : base(sender, e)
        {
            
        }
        /// <summary>
        /// Обработка прайсов, таких как: Каталог OJ 2016_06_01 вер. 6
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
            var needOption = true;
            var baseConnecter = new BaseConnecter(OnBaseMsgAction);
            for (var i = 2; i < row; i++)
            {
                if (Worker.CancellationPending)
                {
                    E.Cancel = true;
                    ResultingPrice.Clear();
                    break;
                }
                if (i == 3) continue;
                var line = new OutputPriceLine();
                var str = ConverterToString(range.Cells[i, 1] as Range);
                if (str.Contains("Рисунок"))
                {
                    // после этого момента опции не читаем
                    needOption = false;
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(str))
                {
                    category1 = str;
                    continue;
                }
                line.Category1 = category1;
                if (line.Category1.Contains("Услуги")) continue; // Если надо, удалим раздел Услуги
                line.VendorCode = ConverterToString(range.Cells[i, 2] as Range);
                var описание = ConverterToString(range.Cells[i, 3] as Range);

                if (string.IsNullOrEmpty(line.VendorCode) && !string.IsNullOrEmpty(описание))
                {
                    // todo: случай когда артикуль не заполнен тоже нужно обработать
                    continue;
                }

                line.Cost = ConverterToString(range.Cells[i, 6] as Range);
                var особенностиУстановки = ConverterToString(range.Cells[i, 11] as Range);
                // todo: вот такое формирование наименование пока под вопросом, нужно выяснить точно как его формировать в автоматическом режиме
                var newname = baseConnecter.OJ_Composition(category1);
                line.Name = string.Format("{0} {1}", newname, line.VendorCode);
                line.Producer = GetManufacturers(line.Name);
                line.ProductDescription = string.Format("<p>{0}</p><p>{1}</p>", описание, особенностиУстановки);
                line.Qt = "1000";
                if (needOption)
                {
                    var opc = ConverterToString(range.Cells[i, 7] as Range);
                    line.Option = OjOptionParser(opc);
                }

                if (string.IsNullOrEmpty(описание) && string.IsNullOrEmpty(str)) break;

                if (!string.IsNullOrEmpty(описание))
                    ResultingPrice.Add(line);
            }
            baseConnecter.Dispose();
        }
        /// <summary>
        /// Поиск опции, прайс Каталог OJ 2016_06_01 вер. 6
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private static string OjOptionParser(string s)
        {
            string input;
            if (string.IsNullOrWhiteSpace(s))
                return string.Empty;
            if (s.Contains("-"))
                return string.Empty;
            if (s.Contains("(опция"))
            {
                input = s.Replace("опция", string.Empty)
                    .Replace(")", string.Empty)
                    .Replace("(", ";")
                    .Replace(",", string.Empty);
                return input;
            }
            input = s.Replace("опция", string.Empty).Trim();
            if (input.Length < 1)
                return string.Empty;
            if (input[0] == '(') input = input.Replace("(", string.Empty);
            input = input.Replace(",", ";").Replace("(", ";").Replace(")", string.Empty);
            return input;
        }

        private void OnBaseMsgAction(string s)
        {
            if (OnMsg != null) OnMsg(s);
        }
    }
}
