using System.Data;
using Soul.Excel;
using Soul.Excel.Test;

var test = new ExcelDocumentTest();


var table = ExcelReader.Read("D:\\报告用纸模板.xlsx", c =>
{
    c.RowIndex = 1;
    c.IsXlsx = true;
});

Console.WriteLine(  );