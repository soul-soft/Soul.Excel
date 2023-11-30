using System;
using System.Collections.Generic;

namespace Soul.Excel
{
    public class ExcelHeaderCollection : List<ExcelDataHeader>
    {
        public void Add(Action<ExcelDataHeader> configure)
        {
            var row = new ExcelDataHeader();
            configure(row);
            Add(row);
        }
    }
}
