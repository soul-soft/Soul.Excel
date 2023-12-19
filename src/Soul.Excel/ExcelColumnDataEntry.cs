namespace Soul.Excel
{
    public class ExcelColumnDataEntry
    {
        public object Data { get; internal set; }
        public int RowSpan { get; set; } = 1;
        public int ColSpan { get; set; } = 1;
       
        public ExcelColumnDataEntry(object name)
        {
            Data = name;
        }

        public ExcelColumnDataEntry(object name, int rowSpan = 1, int colSpan = 1)
        {
            Data = name;
            RowSpan = rowSpan;
            ColSpan = colSpan;
        }

        public override string ToString()
        {
            return Data?.ToString();
        }
    }
}
