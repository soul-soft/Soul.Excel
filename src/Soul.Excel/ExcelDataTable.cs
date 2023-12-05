using System;

namespace Soul.Excel
{
    public class ExcelDataTable
    {
        public string Title { get; set; }

        public string Name { get; set; }

        public ExcelHeaderCollection Headers { get; } = new ExcelHeaderCollection();

        public ExcelHeaderCollection Footers { get; } = new ExcelHeaderCollection();

        public ExcelDataColumnCollection Columns { get; } = new ExcelDataColumnCollection();

        public ExcelDataRowCollection Rows { get; } = new ExcelDataRowCollection();

        public ExcelDataTable()
        {

        }

        public ExcelDataTable(string name) : this(name, name)
        {

        }

        public ExcelDataTable(string name, string title)
        {
            Name = name;
            Title = title;
        }

        public ExcelDataRow NewRow()
        {
            return new ExcelDataRow(this);
        }

        public ExcelDataTable Clone(string name, string title)
        {
            var table = new ExcelDataTable(name, title);
            foreach (var item in Headers)
            {
                table.Headers.Add(item);
            }
            foreach (var item in Columns)
            {
                table.Columns.Add(item);
            }
            return table;
        }

        public ExcelDataTable Clone(string name)
        {
            return Clone(name, name);
        }

        internal int? FreezeRowIndex;
        
        internal int? FreezeColIndex;

        public void Freeze(int rowIndex,int colIndex)
        {
            FreezeRowIndex = rowIndex;
            FreezeColIndex = colIndex;
        }
    }
}
