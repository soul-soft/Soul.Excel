// See https://aka.ms/new-console-template for more information
using System.Data;
using Soul.Excel;
using Soul.Excel.Test;

var test = new ExcelDocumentTest();
test.TestWrite("D:\\wjf.xls");
//test.TestRead("D:\\wjf.xls");
Console.WriteLine("Hello, World!");
