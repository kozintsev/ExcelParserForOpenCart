using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

namespace ExcelParserForOpenCart
{
    public class BaseConnecter : IDisposable
    {
        private readonly SQLiteConnection _connection;
        private readonly bool _isConnected;
        private readonly bool _notExists;
        // делегат для обработки ошибок
        private Action<string> _onMsgAction;

        public BaseConnecter(Action<string> onMsgAction)
        {
            _isConnected = true;
            _onMsgAction = onMsgAction;
            _notExists = false; 
            var databaseName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "base.sqlite");
            if (!File.Exists(databaseName))
            {
                _onMsgAction("Отсутствует файл базы данных");
                _notExists = true;
                return;
            }
            _connection = new SQLiteConnection(string.Format("Data Source={0};", databaseName));
            try
            {
                _connection.Open();
            }
            catch
            {
                _isConnected = false;
                _onMsgAction("Ошибка подключения к базе данных");
            }
        }
        /// <summary>
        /// Метод перезаписывающий наименование категории записанную во множественном числе в единственное число
        /// Например: Бамперы силовые -> Бампер силовой
        /// Служит для формировании наименования продукции
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public string OJ_Composition(string source)
        {
            if (_isConnected == false) return source;
            const string commandText = "SELECT old, new FROM oj_rowscomparsion";
            var myCommand = _connection.CreateCommand();
            myCommand.CommandText = commandText;
            var dataReader = myCommand.ExecuteReader();
            while (dataReader.Read())
            {
                var old = dataReader["old"].ToString();
                old = old.Replace("/n", "").Trim();
                if (!source.Contains(old)) continue;
                var @new = dataReader["new"].ToString();
                return source.Replace(old, @new);
            }
            return source;
        }


        /// <summary>
        /// Возвращает List c производителями
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Manufacturer> GetManufacturers()
        {
            if (_isConnected == false) return null;
            const string commandText = "SELECT id, name, ru_name FROM producers";
            var producers = new List<Manufacturer>();

            var myCommand = _connection.CreateCommand();
            myCommand.CommandText = commandText;
            var dataReader = myCommand.ExecuteReader();
            var i = 1;
            while (dataReader.Read())
            {
                var line = new Manufacturer
                {
                    Id = i,
                    Name = dataReader["name"].ToString(),
                    RuName = dataReader["ru_name"].ToString()
                };
                producers.Add(line);
                i++;
            }
            return producers;
        }

        public void Dispose()
        {
            if (_notExists) return;
            _connection.Close();
            _connection.Dispose();
        }
    }
}
