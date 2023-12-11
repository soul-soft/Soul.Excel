// See https://aka.ms/new-console-template for more information
using System.Data;
using Soul.Excel;
using Soul.Excel.Test;

//var test = new ExcelDocumentTest();
//test.TestWrite("D:\\wjf.xlsx");
//test.TestRead("D:\\wjf.xls");


var table = new ExcelDataTable("标题");
table.Headers.Add(row =>
{
    var firstName = @"填表说明：
1、仅支持数电发票的导入开具，不支持纸质发票批量导入开具；
2、发票流水号：纳税人自定义，长度不超过20位；发票流水号为导入开具区分发票的唯一标识，用于关联发票的明细信息、特定业务信息、附加要素信息；
3、发票类型：增值税专用发票、普通发票；
4、不支持差额开具和减按开具模式；
5、所有内容输入均为文本格式输入；
6、若开具发票需要使用邮箱推送时：可填写购买方电子邮箱；
7、“放弃享受减按1%征收率原因”填写说明：您在2023年1月1日以后取得的适用3%征收率的应税销售收入，可减按1%征收率征收增值税。若您有特殊情况，需要开具其他发票，请在【放弃享受减按1%征收率原因】字段中选择相应原因。
8、含税标志: 在填写sheet页“2-发票明细信息”中,根据实际业务需要,当单价、金额为含税时,选择“是”,当单价、金额为不含税时,选择“否”。
9、受票方自然人标识：如开票给自然人，此标识请选择“是”，购买方名称填写必须大于一个字符。若选择否或为空则意为开票给单位。
10、开具除特定业务外的普通发票，如受票方为自然人，请根据实际需要填写姓名或姓。（例如：张某某，可在名称栏次填写：张某某、张先生或张女士）；如受票方为自然人并要求能将发票归集在个人票夹中展示，请填写姓名及身份证号（自然人纳税人识别号）。如受票方为个体工商户，需填写社会统一信用代码或纳税人识别号，并在受票方自然人标识栏次选择“否”。11、当“特定业务类型”为“农产品收购”、“光伏收购”时,本模板中填写的“购买方名称”指实际的销售方名称，“购买方纳税人识别号”指实际的销售方纳税人识别号。
12、当“特定业务类型”为“农产品收购”时，“购买方名称”、“证件类型”、“购买方纳税人识别号”为必填项。";
    row.Height = 3500;
    row.Add(firstName, c =>
    {
        c.WrapText = true;
        c.Alignment = ExcelAlignment.Left;
        c.ColSpan = 26;
    });
});

table.Headers.Add(row =>
{
    row.Add(@"必填(限20字符)");
    row.Add(@"必填
(限10字符)");
    row.Add(@"非必填
(限10字符)");
    row.Add(@"必填(是/否)
(限2字符)");
    row.Add(@"非必填(是/否)
 (限2字符)");
    row.Add(@"必填
(限100字符)");
    row.Add(@"非必填
(限20字符)");
    row.Add(@"专票必填
(限20字符)");
    row.Add(@"非必填
(限100字符)");
    row.Add(@"非必填
(限100字符)");
    row.Add(@"非必填
(限100字符)");
    row.Add(@"非必填
(限100字符)");
    row.Add(@"非必填
(限230字符)");
    row.Add(@"非必填（是/否）
（限2字符）");
    row.Add(@"非必填
(限100字符)");
    row.Add(@"非必填
(限60字符)");
    row.Add(@"非必填（是/否）
（限2字符）");
    row.Add(@"非必填
(限72字符)");
    row.Add(@"非必填
(限150字符)");
    row.Add(@"非必填
(限40字符)");
    row.Add(@"非必填
(限30字符)");
    row.Add(@"非必填
(限40字符)");
    row.Add(@"非必填
(限20字符)");
    row.Add(@"非必填
(限100字符)");
    row.Add(@"非必填
(限16字符)");
    row.Add(@"非必填
(限16字符)");
});

table.Columns.Add("发票流水号", o =>
{
    o.WrapText = true;
});
table.Columns.Add("发票类型", o =>
{
    o.WrapText = true;
});
table.Columns.Add("特定业务类型", o =>
{
    o.WrapText = true;
});
table.Columns.Add("是否含税", o =>
{
    o.WrapText = true;
});
table.Columns.Add("受票方自然人标识", o =>
{
    o.WrapText = true;
});
table.Columns.Add("购买方名称", o =>
{
    o.WrapText = true;
});
table.Columns.Add("证件类型", o =>
{
    o.WrapText = true;
});
table.Columns.Add("购买方纳税人识别号", o =>
{
    o.WrapText = true;
});
table.Columns.Add("购买方地址", o =>
{
    o.WrapText = true;
});
table.Columns.Add("购买方电话", o =>
{
    o.WrapText = true;
});
table.Columns.Add("购买方开户银行", o =>
{
    o.WrapText = true;
});
table.Columns.Add("购买方银行账号", o =>
{
    o.WrapText = true;
});
table.Columns.Add("备注", o =>
{
    o.WrapText = true;
});
table.Columns.Add("是否展示购买方银行账号", o =>
{
    o.WrapText = true;
});
table.Columns.Add("销售方开户行", o =>
{
    o.WrapText = true;
});
table.Columns.Add("销售方银行账号", o =>
{
    o.WrapText = true;
});
table.Columns.Add("是否展示销售方银行账号", o =>
{
    o.WrapText = true;
});
table.Columns.Add("购买方邮箱", o =>
{
    o.WrapText = true;
});
table.Columns.Add("购买方经办人姓名", o =>
{
    o.WrapText = true;
});
table.Columns.Add("购买方经办人证件类型", o =>
{
    o.WrapText = true;
});
table.Columns.Add("购买方经办人证件号码", o =>
{
    o.WrapText = true;
});
table.Columns.Add("经办人国籍(地区)", o =>
{
    o.WrapText = true;
});
table.Columns.Add("经办人自然人纳税人识别号", o =>
{
    o.WrapText = true;
});
table.Columns.Add("放弃享受减按1%征收率原因", o =>
{
    o.WrapText = true;
});
table.Columns.Add("收款人", o =>
{
    o.WrapText = true;
});
table.Columns.Add("复核人", o =>
{
    o.WrapText = true;
});

ExcelWriter.Write("D:\\wjf.xls", true, table);
