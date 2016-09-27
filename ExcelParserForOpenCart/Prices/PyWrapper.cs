using System;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;
using System.ComponentModel;
using System.IO;
using Microsoft.Office.Interop.Excel;


namespace ExcelParserForOpenCart.Prices
{
    public class PyWrapper : GeneralMethods
    {
        private readonly ScriptEngine _engine;
        private readonly string _folder;
        private const string DetermineFile = "DetermineTypeOfPriceList.py";

        public PyWrapper(object sender, DoWorkEventArgs e)
            : base(sender, e)
        {
            _engine = Python.CreateEngine();
            _folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Analyzers");   
        }

        public void Analyze(Range range)
        {
            var detFile = Path.Combine(_folder, DetermineFile);
            if (!File.Exists(detFile)) 
                return;
            var scope = _engine.CreateScope();
            _engine.ExecuteFile(detFile, scope);
            dynamic determineType = scope.GetVariable("determine_type");
            dynamic result = determineType(range);
            if (result == null) 
                return;
            var fileName = result.ToString();
            if (string.IsNullOrWhiteSpace(fileName)) 
                return;
            var filePath = Path.Combine(_folder, fileName);
            if (!File.Exists(filePath)) 
                return;
            _engine.Execute(filePath, scope);
            dynamic analyze = scope.GetVariable("analyze");
        }
    }
}
