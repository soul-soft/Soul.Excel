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
            foreach (var item in Table.Columns)
            {
                SetValue(item.Name, null);
            }
        }

        public void SetValue(string name, object value, int rowSpan = 1, int colSpan = 1)
        {
            var column = Table.Columns.Where(a => a.Name == name).FirstOrDefault();
            if (column == null)
            {
                throw new InvalidOperationException($"Column not found:\"{name}\"");
            }
            if (_items.ContainsKey(name))
            {
                var item = _items[name];
                item.Data = value;
                item.RowSpan = rowSpan;
                item.ColSpan = colSpan;
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
       
        public int GetInt32(string name)
        {
            var data = this[name];
            if (data == null)
            {
                return 0;
            }
            return Convert.ToInt32(data);
        }

        public long GetInt64(string name)
        {
            var data = this[name];
            if (data == null)
            {
                return 0;
            }
            return Convert.ToInt64(data);
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

        public decimal GetDecimal(string name)
        {
            var data = this[name];
            if (data == null)
            {
                return 0;
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

        public DateTime? GetExactDateTime(string name, string format = "yyyy-MM-dd")
        {
            var data = GetString(name);
            if (data == null)
            {
                return null;
            }
            return DateTime.ParseExact(data, format, System.Globalization.CultureInfo.CurrentCulture);
        }

        public object GetValue(string name)
        {
            return this[name];
        }

        public bool IsNullOrEmpty(string name)
        {
            var value = GetString(name);
            return string.IsNullOrEmpty(value);
        }
    }
}
