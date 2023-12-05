using System;
using NPOI.SS.UserModel;

namespace Soul.Excel
{
    public class ExcelDataColumn
    {
        public string Name { get; }

        public Type DataType { get; set; }

        public int? Width { get; set; } = 8;

        public string DataFormat { get; set; }

        public object DefaultValue { get; set; }

        public bool WrapText { get; set; } = false;

        public ExcelAlignment? Alignment { get; set; }

        public ExcelDataColumn(string name)
        {
            Name = name;
        }
       
        private ICellStyle _dataStyle;

        internal ICellStyle GetDataStyle(IWorkbook book)
        {
            if (_dataStyle == null)
            {
                var cellStyle = book.CreateCellStyle();
                var dataFormat = book.CreateDataFormat();
                if (!string.IsNullOrEmpty(DataFormat))
                {
                    cellStyle.DataFormat = dataFormat.GetFormat(DataFormat);
                }
                if (Alignment == ExcelAlignment.Left)
                {
                    cellStyle.Alignment = HorizontalAlignment.Left;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                }
                else if (Alignment == ExcelAlignment.Right)
                {
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                }
                else
                {
                    cellStyle.Alignment = HorizontalAlignment.Center;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                }
                cellStyle.WrapText = WrapText;
                _dataStyle = cellStyle;
            }
            return _dataStyle;
        }
    }
}
