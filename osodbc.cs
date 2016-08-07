using ScriptEngine.HostedScript.Library.ValueTable;
using ScriptEngine.Machine;
using ScriptEngine.Machine.Contexts;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace osodbc
{
    [ContextClass("ODBC", "ODBC")]
    public class osodbc : AutoContext<osodbc>, IDisposable
    {
        private string _version = "0.0.1";
        [ContextProperty("Версия", "Version")]
        public string Version
        {
            get
            {
                return this._version;
            }
        }

        [ContextProperty("СтрокаПодключения", "ConnectionString")]
        public string ConnectionString { get; set; }
        private IDbConnection dbСonnection;

        [ContextProperty("ТекстЗапроса", "QueryText")]
        public string QueryText { get; set; }

        [ContextProperty("Открыта", "IsOpen")]
        public bool IsOpen { get { if (dbСonnection.State == ConnectionState.Open) return true; else return false; } }

        private Dictionary<string, object> parameters;


        public osodbc()
        {
            parameters = new Dictionary<string, object>();
        }

        [ScriptConstructor(Name = "По умолчанию")]
        public static IRuntimeContextInstance Constructor()
        {
            var osObject = new osodbc();            
            return osObject;
        }

        // Открыть соединение с базой данных
        [ContextMethod("Открыть", "Open")]
        public void Open()
        {           
            dbСonnection = new OdbcConnection(this.ConnectionString);
            dbСonnection.Open();            
        }

        // Закрыть соединение с базой данных
        [ContextMethod("Закрыть", "Close")]
        public void Close()
        {
            if(dbСonnection.State == ConnectionState.Open)
            {
                dbСonnection.Close();
                dbСonnection.Dispose();
            }
            
        }

        // Открыть соединение с базой данных по указанной строке соединения
        [ContextMethod("Открыть", "Open")]
        public void Open(string _connectionString)
        {
            this.ConnectionString = _connectionString;
            this.Open();
        }

        // Установить параметр
        [ContextMethod("УстановитьПараметр", "SetParameters")]
        public void SetParameters(string name, object value)
        {
            parameters.Add(name, value);
        }

        // Выполняет запрос к базе данных
        [ContextMethod("ВыполнитьЗапрос", "RunQuery")]
        public ValueTable RunQuery()
        {
            ValueTable tbl = new ValueTable();
            OdbcCommand dbCommand = (OdbcCommand) dbСonnection.CreateCommand();
            
            // Добавляем параметры в запрос
            string pattern = @"@[A-z0-9]*";
            Regex regex = new Regex(pattern);  
            Match match = regex.Match(QueryText);
            while (match.Success)
            {
                string currParam = match.Value;
                var currValue = parameters[currParam];                             
                dbCommand.Parameters.AddWithValue(currParam,currValue);
                match = match.NextMatch();
            }
            string prepareQuery = regex.Replace(QueryText, "?");
            dbCommand.CommandText = prepareQuery;
            // Выполняем запрос
            IDataReader dbReader = dbCommand.ExecuteReader();
            // Создаем ТаблицуЗначений в которую будем выводить результат запроса
            
            for (int i = 0; i < dbReader.FieldCount; i++)
            {
                string ColumnName = dbReader.GetName(i);
                tbl.Columns.Add(ColumnName);
            }


            while (dbReader.Read())
            {
                for (int i = 0; i < dbReader.FieldCount; i++)
                {
                    string ColumnName = dbReader.GetName(i);
                    ValueTableRow row = tbl.Add();
                    var value = dbReader[ColumnName];
                    System.Type valueType = value.GetType();
                    if(valueType == typeof(int))
                    {
                        row.Set(i, ValueFactory.Create((int)value));
                    }
                    else if (valueType == typeof(Int16))
                    {
                        row.Set(i, ValueFactory.Create((Int16)value));
                    }
                    else if (valueType == typeof(Int32))
                    {
                        row.Set(i, ValueFactory.Create((Int32)value));
                    }
                    else if (valueType == typeof(string))
                    {
                        row.Set(i, ValueFactory.Create((string)value));
                    }
                    else if (valueType == typeof(DateTime))
                    {
                        row.Set(i, ValueFactory.Create((DateTime)value));
                    }


                }
            }
            
            dbReader.Close();
            dbReader = null;
            dbCommand.Dispose();
            dbCommand = null;
            return tbl;

        }


        #region IDisposable Support
        private bool disposedValue = false; // Для определения избыточных вызовов

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: освободить управляемое состояние (управляемые объекты).
                }

                // TODO: освободить неуправляемые ресурсы (неуправляемые объекты) и переопределить ниже метод завершения.
                // TODO: задать большим полям значение NULL.

                disposedValue = true;
            }
        }

        // TODO: переопределить метод завершения, только если Dispose(bool disposing) выше включает код для освобождения неуправляемых ресурсов.
        // ~osodbc() {
        //   // Не изменяйте этот код. Разместите код очистки выше, в методе Dispose(bool disposing).
        //   Dispose(false);
        // }

        // Этот код добавлен для правильной реализации шаблона высвобождаемого класса.
        public void Dispose()
        {
            // Не изменяйте этот код. Разместите код очистки выше, в методе Dispose(bool disposing).
            Dispose(true);
            // TODO: раскомментировать следующую строку, если метод завершения переопределен выше.
            // GC.SuppressFinalize(this);
        }
        #endregion

    }
}
