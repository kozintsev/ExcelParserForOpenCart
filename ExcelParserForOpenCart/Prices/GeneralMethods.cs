using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    /// <summary>
    /// Общий абстрактный класс, который содержит:
    /// 1. BackgroundWorker в котором происходит обработка прайса
    /// 2. Методы для конвертации данных из ячеек
    /// 3. Поиск производителя товара
    /// </summary>
    public abstract class GeneralMethods
    {
        protected readonly BackgroundWorker Worker;
        protected readonly DoWorkEventArgs E;
        protected List<Manufacturer> Manufacturers { get; private set; }

        public List<OutputPriceLine> ResultingPrice { get; private set; }

        protected event Action<string> OnMsg;

        protected GeneralMethods(object sender, DoWorkEventArgs e)
        {
            Worker = sender as BackgroundWorker;
            E = e;
            ResultingPrice = new List<OutputPriceLine>();
            AddManufacturers();
        }

        protected GeneralMethods()
        {
            AddManufacturers();
        }

        protected static string ConverterToString(dynamic obj)
        {
            string s;
            try
            {
                s = Convert.ToString(obj);
            }
            catch
            {
                s = string.Empty;
            }
            return s;
        }

        protected static decimal ConverterToDecimal(Range range)
        {
            if (range == null)
                return 0;
            var obj = range.Value2;
            if (obj == null)
                return 0;
            decimal d;
            try
            {
                d = Convert.ToDecimal(obj);
            }
            catch
            {
                d = 0;
            }
            return d;
        }

        protected static string ConverterToString(Range range)
        {
            if (range == null)
                return string.Empty;
            var obj = range.Value2;
            if (obj == null)
                return string.Empty;
            string s;
            try
            {
                s = Convert.ToString(obj);
            }
            catch
            {
                s = string.Empty;
            }
            return s;
        }
        /// <summary>
        /// Поиск производителя по совпадению в тексте Наименования 
        /// </summary>
        /// <param name="name">Наименование товара в прайс-листе</param>
        /// <returns></returns>
        protected string GetManufacturers(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return string.Empty;
            var tempName = name.ToUpper();
            foreach (var manufacturer in Manufacturers)
            {
                if (!string.IsNullOrWhiteSpace(manufacturer.Name) && tempName.Contains(manufacturer.Name.ToUpper()))
                    return manufacturer.Name;
                //поищем по русским именам
                if (!string.IsNullOrWhiteSpace(manufacturer.RuName) && tempName.Contains(manufacturer.RuName.ToUpper()))
                    return manufacturer.Name;
            }
            return string.Empty;
        }

        private void AddManufacturers()
        {
            Manufacturers = new List<Manufacturer>();
            using (var baseConnecter = new BaseConnecter(OnBaseMsgAction))
            {
                Manufacturers.AddRange(baseConnecter.GetManufacturers());
            }
        }

        private void OnBaseMsgAction(string s)
        {
            if (OnMsg != null) OnMsg(s);
        }
    }
}
