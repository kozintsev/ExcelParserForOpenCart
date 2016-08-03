using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using Microsoft.Office.Interop.Excel;
using ExcelParserForOpenCart.Prices;


namespace ExcelParserForOpenCart
{
    public partial class ExcelParser
    {
        public event Action<string> OnParserAction;
        public event Action<int> OnProgressBarAction;
        public event EventHandler OnOpenDocument;
        public event EventHandler OnSaveDocument;

        private readonly bool _isExcelInstal;
        private BackgroundWorker _workerSave;
        private BackgroundWorker _workerOpen;
        private List<OutputPriceLine> _list;
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
            }
        }

        public bool IsStart()
        {
            if (_workerOpen != null) return _workerOpen.IsBusy;
            if (_workerSave != null) return _workerSave.IsBusy;
            return false;
        }

        public void CancelParsing()
        {
            if (_workerOpen != null && _workerOpen.IsBusy) _workerOpen.CancelAsync();
            if (_workerSave != null && _workerSave.IsBusy) _workerSave.CancelAsync();
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
            _workerOpen = new BackgroundWorker { WorkerReportsProgress = true, WorkerSupportsCancellation = true };
            _workerOpen.DoWork += DoWorkOpen;
            _workerOpen.RunWorkerCompleted += RunCompletedOpenWorker;
            _workerOpen.ProgressChanged += ProgressChangedWorkerOpen;
            _workerOpen.RunWorkerAsync();
        }

        public void SaveResult(string fileName)
        {
            if (_isExcelInstal == false)
                return;
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
            if (_list == null || _list.Count < 1) return;
            _workerSave = new BackgroundWorker { WorkerReportsProgress = true, WorkerSupportsCancellation = true};
            _workerSave.DoWork += DoWorkSave;
            _workerSave.RunWorkerCompleted += RunWorkerCompletedWorkerSave;
            _workerSave.ProgressChanged += ProgressChangedWorkerSave;
            _workerSave.RunWorkerAsync();
        }

        private void ProgressChangedWorkerOpen(object sender, ProgressChangedEventArgs e)
        {
            if (e.UserState != null && !string.IsNullOrEmpty(e.UserState.ToString()))
                SendMessage(e.UserState.ToString());
            SendProgressBarInfo(e.ProgressPercentage);
        }

        private void RunCompletedOpenWorker(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                SendMessage("Отменён анализ документа: " + _openFileName);
            }
            else
            {
                SendMessage("Завершён анализ документа: " + _openFileName);
                if (OnOpenDocument != null) OnOpenDocument(null, null);
            }
        }

        private void DoWorkOpen(object sender, DoWorkEventArgs e)
        {
            _list.Clear();
            _workerOpen.ReportProgress(0);
            var application = new Application();
            var workbook = application.Workbooks.Open(_openFileName);
            var worksheet = workbook.Worksheets[1] as Worksheet;
            if (_workerOpen.CancellationPending)
            {
                application.Quit();
                ReleaseObject(worksheet);
                ReleaseObject(workbook);
                ReleaseObject(application);
                _workerOpen.ReportProgress(0);
                e.Cancel = true;
                return;
            }
            if (worksheet == null) return;
            var range = worksheet.UsedRange;
            var row = worksheet.Rows.Count;
            _workerOpen.ReportProgress(10);
            // todo: метод для определения прайс листа должен вызываться здесь
            PriceType = DetermineTypeOfPriceList(range);
            switch (PriceType)
            {
                case EnumPrices.ДваСоюза:
                    var for2Union = new For2Union(sender, e);
                    for2Union.Analyze(row, range);
                    _list = for2Union.List;
                    break;
                case EnumPrices.OJ:
                    var ojPrice = new OjPrice(sender, e);
                    ojPrice.OnMsg += s =>
                    {
                        _workerOpen.ReportProgress(20, s);
                    };
                    ojPrice.Analyze(row, range);
                    break;
                case EnumPrices.ПТГрупп:
                    break;
                case EnumPrices.Autogur73:
                    var autogurPrice = new AutogurPrice(sender, e);
                    autogurPrice.Analyze(row, range);
                    _list = autogurPrice.List;
                    break;
                case EnumPrices.Композит:
                    break;
                case EnumPrices.Риваль:
                    break;
                case EnumPrices.Неизвестный:
                    _workerOpen.ReportProgress(0, "Прайс не опознан");
                    break;
                default:
                    _workerOpen.ReportProgress(0, "Прайс не опознан");
                    break;
            }
            application.Quit();
            ReleaseObject(worksheet);
            ReleaseObject(workbook);
            ReleaseObject(application);
            _workerOpen.ReportProgress(!e.Cancel ? 50 : 0);
        }

        private void ProgressChangedWorkerSave(object sender, ProgressChangedEventArgs e)
        {
            SendProgressBarInfo(e.ProgressPercentage);
        }

        private void RunWorkerCompletedWorkerSave(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                SendMessage("Отменено сохранение документа: " + _saveFileName);
            }
            else
            {
                SendMessage("Прайс создан! Сохраняю как: " + _saveFileName);
                if (OnSaveDocument != null) OnSaveDocument(null, null);   
            }
        }

        private void DoWorkSave(object sender, DoWorkEventArgs e)
        {
            _workerSave.ReportProgress(65);
            var application = new Application();
            var workbook = application.Workbooks.Open(_template);
            var worksheet = workbook.Worksheets[1] as Worksheet;
            if (worksheet == null) return;
            _workerSave.ReportProgress(70);
            if (_workerSave.CancellationPending)
            {
                application.Quit();
                ReleaseObject(worksheet);
                ReleaseObject(workbook);
                ReleaseObject(application);
                _workerSave.ReportProgress(50);
                e.Cancel = true;
                return;
            }
            // действия по заполнению шаблона
            var i = 2;
            foreach (var obj in _list)
            {
                if (_workerSave.CancellationPending)
                {
                    e.Cancel = true;
                    break;
                }
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
            if (!_workerSave.CancellationPending) worksheet.SaveAs(_saveFileName);
            application.Quit();
            ReleaseObject(worksheet);
            ReleaseObject(workbook);
            ReleaseObject(application);
            _workerSave.ReportProgress(!e.Cancel ? 100 : 50);
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
