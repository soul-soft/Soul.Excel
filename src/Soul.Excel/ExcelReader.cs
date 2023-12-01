using System;
using System.IO;

namespace Soul.Excel
{
    public static class ExcelReader
    {
        public static ExcelDataTable Read(string file, Action<ExcelReaderOptions> configure)
        {
            return new ExcelDocument().Read(file, configure);
        }

        public static ExcelDataTable Read(Stream stream, Action<ExcelReaderOptions> configure)
        {
            return new ExcelDocument().Read(stream, configure);
        }
    }
}
