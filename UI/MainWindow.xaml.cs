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
            _excelParser = new ExcelParser();
            InitializeComponent();
            _excelParser.OnParserAction += OnParserAction;
            _excelParser.OnProgressBarAction += OnProgressBarAction;
            _excelParser.OnOpenDocument += OnOpenDocument;
        }

        private void OnOpenDocument(object sender, System.EventArgs e)
        {
            BtnSave.IsEnabled = true;            
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
            _excelParser.OpenExcel(filename);
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
            _excelParser.SaveResult(filename);
        }

        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            switch (ComBoxSelectPrice.SelectedIndex)
            {
                case (int)EnumPrices.ДваСоюза:

                    MessageBox.Show("0");
                    _excelParser.PriceType = EnumPrices.ДваСоюза;
                    break;
                case (int)EnumPrices.OJ:

                    MessageBox.Show("1");
                    _excelParser.PriceType = EnumPrices.OJ;
                    break;
                case (int)EnumPrices.tdgroup:
                    MessageBox.Show("2");
                    _excelParser.PriceType = EnumPrices.tdgroup;
                    break;
                case (int)EnumPrices.lapter:
                    MessageBox.Show("3");
                    _excelParser.PriceType = EnumPrices.lapter;
                    break;
                case (int)EnumPrices.composite:
                    MessageBox.Show("4");
                    _excelParser.PriceType = EnumPrices.composite;
                    break;
                case (int)EnumPrices.rival:
                    MessageBox.Show("5");
                    _excelParser.PriceType = EnumPrices.rival;
                    break;
                case (int)EnumPrices.ptgroup:
                    MessageBox.Show("6");
                    _excelParser.PriceType = EnumPrices.ptgroup;
                    break;
                case (int)EnumPrices.pyanov:
                    MessageBox.Show("7");
                    _excelParser.PriceType = EnumPrices.pyanov;
                    break;

            }
        }
    }
}
