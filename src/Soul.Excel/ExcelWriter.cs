using System.IO;

namespace Soul.Excel
{
    public static class ExcelWriter
    {
        public static void Write(string file, bool isXlsx = false, params ExcelDataTable[] tables)
        {
            var document = new ExcelDocument();
            document.Tables.AddRange(tables);
            document.Wirte(file, isXlsx);
        }

        public static void Write(Stream stream, bool isXlsx = false, params ExcelDataTable[] tables)
        {
            var document = new ExcelDocument();
            document.Tables.AddRange(tables);
            document.Wirte(stream, true);
        }
    }
}
