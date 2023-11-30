using System;
using System.Collections.Generic;
using System.Linq;

namespace Soul.Excel
{
    public class ExcelDataRow
    {
        public ExcelDataTable Table { get; }

        private Dictionary<string, ExcelDataInfo<object>> _items = new Dictionary<string, ExcelDataInfo<object>>();

        public object[] ItemArray => _items.Values.Select(s => s.Data).ToArray();

        internal ExcelDataRow(ExcelDataTable table)
        {
            Table = table;
        }

        public void SetValue(string name, object value, int rowSpan = 1, int colSpan = 1)
        {
            var column = Table.Columns.Where(a => a.Name == name).FirstOrDefault();
            if (column == null)
            {
                throw new InvalidOperationException($"字段‘{name}’不存在");
            }
            if (_items.ContainsKey(name))
            {
                _items[name] = new ExcelDataInfo<object>(value);
            }
            else
            {
                _items.Add(name, new ExcelDataInfo<object>(value, rowSpan, colSpan));
            }
        }

        internal ExcelDataInfo<object> GetDataInfo(string name)
        {
            return _items[name];
        }

        public object this[string name]
        {
            get
            {
                return _items[name].Data;
            }
            set
            {
                SetValue(name, value);
            }
        }

        public string GetString(string name)
        {
            var data = this[name];
            if (data == null)
            {
                return null;
            }
            return data.ToString();
        }

        public double GetDouble(string name)
        {
            var data = this[name];
            if (data == null)
            {
                return 0;
            }
            return Convert.ToDouble(data);
        }
    
        public double? GetDoubleNullable(string name)
        {
            var data = this[name];
            if (data == null)
            {
                return null;
            }
            return Convert.ToDouble(data);
        }

        public decimal GetDecimal(string name)
        {
            var data = this[name];
            if (data == null)
            {
                return 0;
            }
            return Convert.ToDecimal(data);
        }

        public decimal? GetDecimalNullable(string name)
        {
            var data = this[name];
            if (data == null)
            {
                return null;
            }
            return Convert.ToDecimal(data);
        }

        public DateTime? GetDateTime(string name)
        {
            var data = this[name];
            if (data == null)
            {
                return null;
            }
            return (DateTime?)data;
        }

        public DateTime? ParseExactDateTime(string name, string format = "yyyy-MM-dd HH:mm:ss")
        {
            var data = GetString(name);
            if (data == null)
            {
                return null;
            }
            return DateTime.ParseExact(data, format, System.Globalization.CultureInfo.CurrentCulture);
        }
    }
}
