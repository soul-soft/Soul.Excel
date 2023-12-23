using System.Data;
using Soul.Excel;
using Soul.Excel.Test;

var test = new ExcelDocumentTest();


var table = ExcelReader.Read("E:\\报告用纸导入模板.xls", c =>
{
    c.RowIndex = 1;
    c.IsXlsx = false;
});

Console.WriteLine(  );