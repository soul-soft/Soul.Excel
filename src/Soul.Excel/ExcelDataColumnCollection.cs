using System;
using System.Collections.Generic;

namespace Soul.Excel
{
    public class ExcelDataColumnCollection : List<ExcelDataColumn>
    {
        public void Add(string name)
        {
            var column = new ExcelDataColumn(name);
            Add(column);
        }

        public void Add(string name,Action<ExcelDataColumn> configure)
        {
            var column = new ExcelDataColumn(name);
            configure(column);
            Add(column);
        }
    }
}
