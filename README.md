# Soul.Excel

一款强大的Excel操作工具

``` C#

internal class ExcelDocumentTest
{
    public void TestRead(string file)
    {
        ExcelDocument document = new ExcelDocument();
        var table = document.Read(file, c =>
        {
            c.RowIndex = 3;
        });
        foreach (var row in table.Rows)
        {
            var name = row.GetString("姓名");
            var birthday = row.GetDateTime("生日");
            var gender = row.GetString("性别");
            var chinese = row.GetDouble("语文");
            var english = row.GetDouble("英语");
            var math = row.GetDouble("数学");
            var physics = row.GetDouble("物理");
        }
    }

    public void TestWrite(string file)
    {
        var document = new ExcelDocument();
        var table = new ExcelDataTable("学生信息");
        table.Freeze(4, 3);
        table.Headers.Add(row =>
        {
            row.Add("基本信息", rowSpan: 2, colSpan: 3);
            row.Add("成绩信息", colSpan: 4);
        });
        table.Headers.Add(row =>
        {
            //合并单元格
            row.Add("基本信息", colSpan: 3);
            row.Add("文科成绩", colSpan: 2);
            row.Add("理科成绩", colSpan: 2);
        });
        table.Columns.Add("姓名", o =>
        {
            o.Alignment = ExcelAlignment.Left;
            //自动换行
            o.WrapText = true;
        });
        table.Columns.Add("生日", o =>
        {
            //固定宽度
            o.Width = 8;
            o.DataFormat = "yyyy-MM-dd";
        });
        table.Columns.Add("性别");
        table.Columns.Add("语文", o =>
        {
            //默认值
            o.DefaultValue = 0;
            //数据格式化
            o.DataFormat = "#.00";
            //对齐
            o.Alignment = ExcelAlignment.Right;
        });
        table.Columns.Add("英语", o =>
        {
            o.DataFormat = "#.00";
            o.Alignment = ExcelAlignment.Right;
        });
        table.Columns.Add("数学", o =>
        {
            o.DataFormat = ".00";
            o.Alignment = ExcelAlignment.Right;
        });
        table.Columns.Add("物理", o =>
        {
            o.DataFormat = "#.00";
            o.Alignment = ExcelAlignment.Right;
        });
        //支持Clone
        var table1 = table.Clone("学生信息2");
        decimal? english = new decimal(70.2123);
        for (int i = 0; i < 10; i++)
        {
            var row = table.NewRow();
            row["生日"] = DateTime.Now;
            row["姓名"] = "张三";
            row["性别"] = true;
            row["语文"] = 80.787878;
            row["英语"] = 80;
            row["数学"] = 90;
            row["物理"] = 100;
            table.Rows.Add(row);
        }
        //支持多Sheet
        document.Tables.Add(table);
        document.Tables.Add(table1);
        using (var fs = new FileStream(file, FileMode.OpenOrCreate))
        {
            ExcelWriter.Wirte(fs, true, table, table1);
        }
    }
}

```
