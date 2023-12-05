namespace Soul.Excel
{
    public class ExcelDataInfo<T>
    {
        public T Data { get; internal set; }
        public int RowSpan { get; set; } = 1;
        public int ColSpan { get; set; } = 1;

        public ExcelDataInfo(T name)
        {
            Data = name;
        }

        public ExcelDataInfo(T name, int rowSpan = 1, int colSpan = 1)
        {
            Data = name;
            RowSpan = rowSpan;
            ColSpan = colSpan;
        }
    }
}
