using System.Collections.Generic;

namespace Soul.Excel
{
    public class ExcelDataHeader
    {
        public List<ExcelDataInfo<string>> Items { get; } = new List<ExcelDataInfo<string>>();

        public void Add(string name, int rowSpan = 1, int colSpan = 1)
        {
            Items.Add(new ExcelDataInfo<string>(name, rowSpan, colSpan));
        }
    }
}
