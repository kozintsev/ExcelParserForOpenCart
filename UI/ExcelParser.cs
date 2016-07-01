using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace ExcelParserForOpenCart
{
    class ExcelParser
    {
        public event Action<string> OnParserAction;
        public event Action<int> OnProgressBarAction;
        public event EventHandler OnOpenDocument;

        private readonly bool _isExcelInstal;
        private readonly BackgroundWorker _workerSave;
        private readonly BackgroundWorker _workerOpen;
        private readonly List<OutputPriceLine> _list;
        private string _template;
        private string _openFileName;
        private string _saveFileName;

        public EnumPrices PriceType { get; set; }

        public ExcelParser()
        {
            _isExcelInstal = true;
            _list = new List<OutputPriceLine>();
            if (!IsExcelInstall())
            {
                SendMessage("Excel не установлен!");
                _isExcelInstal = false;
                return;
            }
            _workerSave = new BackgroundWorker { WorkerReportsProgress = true };
            _workerSave.DoWork += _workerSave_DoWork;
            _workerSave.RunWorkerCompleted += _workerSave_RunWorkerCompleted;
            _workerSave.ProgressChanged += _workerSave_ProgressChanged;

            _workerOpen = new BackgroundWorker { WorkerReportsProgress = true };
            _workerOpen.DoWork += _workerOpen_DoWork;
            _workerOpen.RunWorkerCompleted += _workerOpen_RunWorkerCompleted;
            _workerOpen.ProgressChanged += _workerOpen_ProgressChanged;
        }

        public void OpenExcel(string fileName)
        {
            _openFileName = fileName;
            if (_isExcelInstal == false)
                return;
            if (!File.Exists(fileName))
            {
                SendMessage("Ошибка! Файл отсутствует!");
                return;
            }
            _list.Clear();
            _workerOpen.RunWorkerAsync();
        }

        public void SaveResult(string fileName)
        {
            if (_isExcelInstal == false)
                return;
            #region Test
            //var list = new List<OutputPriceLine>();
            //var line = new OutputPriceLine
            //{
            //    Name = "Имя",
            //    VendorCode = "1234567",
            //    Category1 = "Багажники на Suzuki",
            //    Category2 = "Багажники",
            //    ProductDescription = "Огромный багажник, вместительностью до 300 кг",
            //    Cost = "3000",
            //    Foto = "http://img.yandex.ru/2034.jpg",
            //    Option = "Да",
            //    Qt = "1000",
            //    PlusThePrice = "100"
            //};
            //list.Add(line);
            //var line2 = new OutputPriceLine
            //{
            //    Name = "Имя",
            //    VendorCode = "7654321",
            //    Category1 = "Багажники на Suzuki",
            //    Category2 = "Багажники",
            //    ProductDescription = "Огромный багажник, вместительностью до 300 кг",
            //    Cost = "3000",
            //    Foto = "http://img.yandex.ru/2034.jpg",
            //    Option = "Нет",
            //    Qt = "2000",
            //    PlusThePrice = "200"
            //};
            //list.Add(line2);
            //_list = list;
            #endregion
            _saveFileName = fileName;
            _template = Global.GetTemplate();
            if (_template == null)
            {
                SendMessage("Ошибка! Не могу получить путь к шаблону!");
                return;
            }
            if (!File.Exists(_template))
            {
                SendMessage("Ошибка! Отсутствует шаблон!");
                return;
            }
            if (_list != null && _list.Count >= 1)
                _workerSave.RunWorkerAsync();
        }

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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        private static string ConverterToString(dynamic obj)
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

        private void _workerOpen_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            SendProgressBarInfo(e.ProgressPercentage);
        }

        private void _workerOpen_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SendMessage("Завершён анализ документа: " + _openFileName);
            if (OnOpenDocument != null) OnOpenDocument(null, null);
        }

        /// <summary>
        /// Обработка для прайса 2 союза
        /// </summary>
        /// <param name="row"></param>
        /// <param name="range"></param>
        private void For2Union(int row, Range range)
        {
            var category1 = string.Empty;
            var category2 = string.Empty;
            for (var i = 9; i < row; i++)
            {
                // todo: алгоритм требует большой доработки
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
                theRange = range.Cells[i, 3] as Range;
                if (theRange != null)
                {
                    line.VendorCode = ConverterToString(theRange.Value2);
                }
                theRange = range.Cells[i, 4] as Range;
                if (theRange != null)
                    line.Name = ConverterToString(theRange.Value2);
                theRange = range.Cells[i, 5] as Range;
                if (theRange != null)
                    line.Qt = ConverterToString(theRange.Value2);
                theRange = range.Cells[i, 6] as Range;
                if (theRange != null)
                    line.Cost = ConverterToString(theRange.Value2);
                if (!string.IsNullOrEmpty(line.VendorCode))
                    _list.Add(line);
                if (string.IsNullOrEmpty(str)) break;
            }
        }

        private void OjPrice(int row, Range range)
        { }

        private void TdgroupPrice(int row, Range range)
        { }

        private void LapterPrice(int row, Range range)
        { }

        private void CompositePrice(int row, Range range)
        { }

        private void RivalPrice(int row, Range range)
        { }

        private void PtgroupPrice(int row, Range range)
        { }

        private void PyanovPrice(int row, Range range)
        { }

        private void _workerOpen_DoWork(object sender, DoWorkEventArgs e)
        {
            _list.Clear();
            var application = new Application();
            var workbook = application.Workbooks.Open(_openFileName);
            var worksheet = workbook.Worksheets[1] as Worksheet;
            if (worksheet == null) return;
            var range = worksheet.UsedRange;
            var row = worksheet.Rows.Count;
            switch (PriceType)
            {
                case EnumPrices.ДваСоюза:
                    For2Union(row, range);
                    break;
                case EnumPrices.OJ:
                    break;
                default:
                    For2Union(row, range);
                    break;
            }
            application.Quit();
            ReleaseObject(worksheet);
            ReleaseObject(workbook);
            ReleaseObject(application);
            _workerOpen.ReportProgress(50);
        }

        private void _workerSave_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            SendProgressBarInfo(e.ProgressPercentage);
        }

        private void _workerSave_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SendMessage("Прайс создан! Сохраняю как: " + _saveFileName);
        }

        private void _workerSave_DoWork(object sender, DoWorkEventArgs e)
        {
            var application = new Application();
            var workbook = application.Workbooks.Open(_template);
            var worksheet = workbook.Worksheets[1] as Worksheet;
            if (worksheet == null) return;
            // действия по заполнению шаблона
            var i = 2;
            foreach (var obj in _list)
            {
                // заносить полученную линию в шаблон
                worksheet.Cells[i, 1] = obj.VendorCode;
                worksheet.Cells[i, 2] = obj.Name;
                worksheet.Cells[i, 3] = obj.Category1;
                worksheet.Cells[i, 4] = obj.Category2;
                worksheet.Cells[i, 5] = obj.ProductDescription;
                worksheet.Cells[i, 6] = obj.Cost;
                worksheet.Cells[i, 7] = obj.Foto;
                worksheet.Cells[i, 8] = obj.Option;
                worksheet.Cells[i, 9] = obj.Qt;
                worksheet.Cells[i, 10] = obj.PlusThePrice;
                i++;
            }
            worksheet.SaveAs(_saveFileName);
            application.Quit();
            ReleaseObject(worksheet);
            ReleaseObject(workbook);
            ReleaseObject(application);
            _workerSave.ReportProgress(100);
        }



        private void SendMessage(string message)
        {
            if (OnParserAction != null) OnParserAction(message);
        }

        private void SendProgressBarInfo(int i)
        {
            if (OnProgressBarAction != null) OnProgressBarAction(i);
        }

    }
}
