using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace Soul.Excel
{
    public class ExcelDataHeader
    {
        public short? Height { get; set; }

        public List<ExcelDataHeaderEntry> Items { get; } = new List<ExcelDataHeaderEntry>();

        public void Add(string name, int rowSpan = 1, int colSpan = 1)
        {
            Items.Add(new ExcelDataHeaderEntry(name, rowSpan, colSpan));
        }

        public void Add(string name, Action<ExcelDataHeaderEntry> configure)
        {
            var entry = new ExcelDataHeaderEntry(name);
            configure(entry);
            Items.Add(entry);
        }
    }

    public class ExcelDataHeaderEntry
    {
        public string Name { get; set; }
        public int RowSpan { get; set; } = 1;
        public int ColSpan { get; set; } = 1;
        public bool WrapText { get; set; } = true;
        public ExcelAlignment? Alignment { get; set; }

        public ExcelDataHeaderEntry(string name)
        {
            Name = name;
        }

        public ExcelDataHeaderEntry(string name, int rowSpan, int colSpan)
        {
            Name = name;
            RowSpan = rowSpan;
            ColSpan = colSpan;
        }

        private ICellStyle _style;

        internal ICellStyle GetStyle(IWorkbook book)
        {
            if (_style == null)
            {
                var style = book.CreateCellStyle();

                if (Alignment == ExcelAlignment.Left)
                {
                    style.Alignment = HorizontalAlignment.Left;
                    style.VerticalAlignment = VerticalAlignment.Center;
                }
                else if (Alignment == ExcelAlignment.Right)
                {
                    style.Alignment = HorizontalAlignment.Right;
                    style.VerticalAlignment = VerticalAlignment.Center;
                }
                else
                {
                    style.Alignment = HorizontalAlignment.Center;
                    style.VerticalAlignment = VerticalAlignment.Center;
                }
                style.WrapText = WrapText;
                var font = book.CreateFont();
                font.IsBold = true;
                font.FontHeight = 350;
                style.SetFont(font);
                _style = style;
            }
            return _style;
        }
    }
}
