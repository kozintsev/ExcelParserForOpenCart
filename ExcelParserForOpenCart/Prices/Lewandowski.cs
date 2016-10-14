using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    public class Lewandowski : GeneralMethods
    {
        public Lewandowski(object sender, DoWorkEventArgs e)
            : base(sender, e)
        {
            
        }
        /// <summary>
        /// Композит ИП Левандовская И.Л.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="range"></param>
        public void Analyze(int row, Range range)
        {
            if (Worker.CancellationPending)
            {
                E.Cancel = true;
                return;
            }
            var startTable = false;
            var category1 = string.Empty;
            var j = 0;
            ResultingPrice.Clear();
            const string pattern = "[0-9]+";
            for (var i = 7; i < row; i++)
            {
                if (Worker.CancellationPending)
                {
                    E.Cancel = true;
                    ResultingPrice.Clear();
                    break;
                }
                var str = ConverterToString(range.Cells[i, 1] as Range);
                var name = ConverterToString(range.Cells[i, 2] as Range);
                if (string.IsNullOrWhiteSpace(name)) break;
                if (string.IsNullOrWhiteSpace(str.Trim()))
                {
                    startTable = false;
                    category1 = name.Replace(':', ' ').Trim();
                    continue;
                }
                if (str.Contains("№"))
                {
                    startTable = true;
                    j++;
                    continue;
                }
                if (!startTable) continue;
                var postfix = Regex.Match(category1, pattern).Value;
                var cost1 = ConverterToDecimal(range.Cells[i, "F"] as Range);
                var cost2 = ConverterToDecimal(range.Cells[i, "I"] as Range);
                var plus = "";
                var option = "";
                var cost = (cost1 == 0 ? cost2 : cost1).ToString(CultureInfo.CurrentCulture);
                var desk = cost1 == 0 ? "АБС-пл." : "стеклопл.";
                if (cost1 > 0 && cost2 > 0)
                {
                    cost = cost2.ToString(CultureInfo.CurrentCulture);
                    plus = (cost1 - cost2).ToString(CultureInfo.CurrentCulture);
                    option = "стеклопл.";
                    desk = "АБС-пл.";
                }
                var line = new OutputPriceLine
                {
                    VendorCode = string.Format("A-{0}-{1}-{2}", j, str, string.IsNullOrWhiteSpace(postfix) ? category1 : postfix),
                    Name = name.Trim(),
                    Category1 = category1,
                    Cost = cost,
                    PlusThePrice = plus,
                    Option = option,
                    ProductDescription = desk,
                    Qt = "1000"
                };
                line.Producer = GetManufacturers(line.Name);

                if (!string.IsNullOrEmpty(line.Name))
                    ResultingPrice.Add(line);
            }
        }

    }
}
