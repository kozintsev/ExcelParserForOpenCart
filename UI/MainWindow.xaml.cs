using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using System.Windows;
using Microsoft.Win32;

namespace ExcelParserForOpenCart
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private readonly ExcelParser _excelParser;

        public MainWindow()
        {
            InitializeComponent();
            BtnCancel.IsEnabled = false;
            var strVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion;
            Title = string.Format("Конвертер прайслистов (версия: {0})", strVersion);
            _excelParser = new ExcelParser();
            _excelParser.OnParserAction += OnParserAction;
            _excelParser.OnProgressBarAction += OnProgressBarAction;
            _excelParser.OnOpenDocument += OnOpenDocument;
            _excelParser.OnSaveDocument += OnSaveDocument;
        }

        private void OnSaveDocument(object sender, EventArgs eventArgs)
        {
            BtnOpen.IsEnabled = true;
            BtnSave.IsEnabled = false;
            BtnCancel.IsEnabled = false;
        }

        private void OnOpenDocument(object sender, EventArgs e)
        {
            BtnOpen.IsEnabled = true;
            BtnSave.IsEnabled = true;
            BtnCancel.IsEnabled = false;
        }

        private void OnProgressBarAction(int obj)
        {
            ProgressBar.Value = obj;
        }

        private void OnParserAction(string message)
        {
            MessageList.Items.Add(message);
        }

        private string CreateOpenFileDialog()
        {
            var filename = string.Empty;
            var dlg = new OpenFileDialog { Filter = "Excel files|;*.xlsx;*.xls" };
            dlg.FileOk += delegate
            {
                filename = dlg.FileName;
            };
            dlg.ShowDialog(this);
            return filename;
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var filename = CreateOpenFileDialog();
            if (string.IsNullOrEmpty(filename)) return;
            MessageList.Items.Clear();
            MessageList.Items.Add("Открываю и обрабатываю документ.");
            _excelParser.OpenExcel(filename);
            MessageList.Items.Add("Пожалуйста, подождите...");
            BtnOpen.IsEnabled = false;
            BtnSave.IsEnabled = false;
            BtnCancel.IsEnabled = true;
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            var filename = string.Empty;
            var dlg = new SaveFileDialog
            {
                Filter = "Excel files|*.xls"
            };
            dlg.FileOk += delegate
            {
                filename = dlg.FileName;
            };
            dlg.ShowDialog(this);
            if (string.IsNullOrEmpty(filename)) return;
            BtnOpen.IsEnabled = false;
            BtnSave.IsEnabled = false;
            BtnCancel.IsEnabled = true;
            MessageList.Items.Add("Идёт сохранение документа.");
            _excelParser.SaveResult(filename);
            MessageList.Items.Add("Пожалуйста, подождите...");
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!_excelParser.IsStart()) return;
            var result = MessageBox.Show("Идёт работа. Вы увререны что хотите завершить работу?", "Вопрос?",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.No)
                e.Cancel = true;
            //todo: возможна проблема, не выгрузится процесс Excel
            // пофиксено, но нужно тестировать
            if (result == MessageBoxResult.Yes)
                _excelParser.CancelParsing();
            Thread.Sleep(2000);
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            _excelParser.CancelParsing();
        }
    }
}
