﻿using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Soul.Excel
{
    public class ExcelDocument
    {
        public List<ExcelDataTable> Tables { get; } = new List<ExcelDataTable>();

        #region Wirte
        private IWorkbook OpenWorkbook(bool isXlsx, Stream stream)
        {
            if (isXlsx)
            {
                return new XSSFWorkbook(stream);
            }
            return new HSSFWorkbook(stream);
        }

        public void Wirte(string file, bool isXlsx = false)
        {
            using (var fs = new FileStream(file, FileMode.Create))
            {
                Wirte(fs, isXlsx);
            }
        }

        public void Wirte(Stream stream, bool isXlsx = false)
        {
            IWorkbook document;
            if (isXlsx)
            {
                document = new XSSFWorkbook();
            }
            else
            {
                document = new HSSFWorkbook();
            }
            var defaultStyles = new DefaultExcelStyles(document);
            foreach (var item in Tables)
            {
                ISheet sheet;
                if (!string.IsNullOrEmpty(item.Name))
                {
                    sheet = document.CreateSheet(item.Name);
                }
                else
                {
                    sheet = document.CreateSheet();
                }
                WriteTable(item, sheet, defaultStyles);
                if (item.FreezeColIndex > 0 && item.FreezeRowIndex > 0)
                {
                    sheet.CreateFreezePane(item.FreezeColIndex.Value, item.FreezeRowIndex.Value);
                }
            }
            document.Write(stream);
        }

        private void WriteTable(ExcelDataTable table, ISheet sheet, DefaultExcelStyles styles)
        {
            var rangeAddresses = new List<CellRangeAddress>();
            var rowIndex = sheet.PhysicalNumberOfRows;
            if (!string.IsNullOrEmpty(table.Title))
            {
                var row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    var cell = row.CreateCell(row.PhysicalNumberOfCells);
                    cell.SetCellValue(table.Title);
                    cell.CellStyle = styles.TitleStyle;
                }
                AddMergedRegion(sheet, rangeAddresses, rowIndex, rowIndex, 0, table.Columns.Count - 1);
                rowIndex++;
            }
            foreach (var header in table.Headers)
            {
                var cellIndex = 0;
                var row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
                if (header.Height != null)
                {
                    row.Height = header.Height.Value;
                }
                foreach (var item in header.Items)
                {
                    for (var i = cellIndex; i < cellIndex + item.ColSpan; i++)
                    {
                        var cell = row.CreateCell(i);
                        cell.SetCellValue(item.Name);
                        var style = item.GetStyle(sheet.Workbook);
                        DefaultExcelStyles.InitStyle(style);
                        cell.CellStyle = style;
                    }
                    if (item.ColSpan > 1 || item.RowSpan > 1)
                    {
                        AddMergedRegion(sheet, rangeAddresses, rowIndex, rowIndex + item.RowSpan - 1, cellIndex, cellIndex + item.ColSpan - 1);
                    }
                    cellIndex += item.ColSpan;
                }
                rowIndex++;
            }
            if (table.Columns.Any())
            {
                var row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
                var cellIndex = 0;
                foreach (var item in table.Columns)
                {
                    var cell = row.CreateCell(row.PhysicalNumberOfCells);
                    cell.SetCellValue(item.Name);
                    cell.CellStyle = styles.ColumnStyle;
                    if (item.Width != null)
                    {
                        cell.Sheet.SetColumnWidth(cellIndex, item.Width.Value * 500);
                    }
                    cellIndex++;
                }
                rowIndex++;
            }
            foreach (var dataRow in table.Rows)
            {
                var cellIndex = 0;
                var row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
                foreach (var column in table.Columns)
                {
                    var data = dataRow.GetDataInfo(column.Name);
                    for (var i = cellIndex; i < table.Columns.Count; i++)
                    {
                        var cell = row.CreateCell(i);
                        SetCellValue(column, cell, data.Data);
                        var style = column.GetStyle(sheet.Workbook);
                        DefaultExcelStyles.InitStyle(style);
                        cell.CellStyle = style;
                    }
                    if (data.ColSpan > 1 || data.RowSpan > 1)
                    {
                        AddMergedRegion(sheet, rangeAddresses, rowIndex, rowIndex + data.RowSpan - 1, cellIndex, cellIndex + data.ColSpan - 1);
                    }
                    cellIndex += data.ColSpan;
                }
                rowIndex++;
            }
            foreach (var header in table.Footers)
            {
                var cellIndex = 0;
                var row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
                foreach (var item in header.Items)
                {
                    for (var i = cellIndex; i < cellIndex + item.ColSpan; i++)
                    {
                        var cell = row.CreateCell(i);
                        cell.SetCellValue(item.Name);
                        var style = item.GetStyle(sheet.Workbook);
                        DefaultExcelStyles.InitStyle(style);
                        cell.CellStyle = style;
                    }
                    if (item.ColSpan > 1 || item.RowSpan > 1)
                    {
                        AddMergedRegion(sheet, rangeAddresses, rowIndex, rowIndex + item.RowSpan - 1, cellIndex, cellIndex + item.ColSpan - 1);
                    }
                    cellIndex += item.ColSpan;
                }
                rowIndex++;
            }
        }

        private void AddMergedRegion(ISheet sheet, List<CellRangeAddress> rangeAddresses,int firstRow, int lastRow, int firstCol, int lastCol)
        {
            var rangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            if (!rangeAddresses.Any(a => a.FirstColumn <= rangeAddress.FirstColumn && a.LastColumn >= rangeAddress.LastColumn && a.FirstRow <= rangeAddress.FirstRow && a.LastRow >= rangeAddress.LastRow))
            {
                rangeAddresses.Add(rangeAddress);
                sheet.AddMergedRegion(rangeAddress);
            }
        }

        private void SetCellValue(ExcelDataColumn column, ICell cell, object value)
        {
            if (value == null && column.DefaultValue != null)
            {
                value = column.DefaultValue;
            }
            if (value == null)
            {
                cell.SetCellValue(string.Empty);
            }
            else if (value.GetType() == typeof(DateTime?))
            {
                var data = (DateTime?)value;
                cell.SetCellValue(data.Value);
            }
            else if (value.GetType() == typeof(DateTime))
            {
                var data = (DateTime)value;
                cell.SetCellValue(data);
            }
            else if (value.GetType() == typeof(bool))
            {
                var data = (bool)value;
                cell.SetCellValue(data);
            }
            else if (value.GetType() == typeof(bool?))
            {
                var data = (bool?)value;
                cell.SetCellValue(data.Value);
            }
            else if (value.GetType().IsValueType)
            {
                cell.SetCellValue(Convert.ToDouble(value));
            }
            else
            {
                var dataFormat = value.ToString();
                cell.SetCellValue(dataFormat);
            }
        }

        #endregion

        #region Read
        public ExcelDataTable Read(string file, Action<ExcelReaderOptions> configure)
        {
            using (var fs = new FileStream(file, FileMode.Open))
            {
                return Read(fs, configure);
            }
        }
        public ExcelDataTable Read(Stream stream, Action<ExcelReaderOptions> configure)
        {
            var options = new ExcelReaderOptions();
            configure(options);
            IWorkbook document = OpenWorkbook(options.IsXlsx, stream);
            var table = new ExcelDataTable();
            var sheet = document.GetSheetAt(options.SheetIndex);
            var columnRow = sheet.GetRow(options.RowIndex);
            for (int i = 0; i < columnRow.Cells.Count; i++)
            {
                var name = columnRow.GetCell(i);
                table.Columns.Add(name.StringCellValue);
            }
            for (int i = options.RowIndex + 1; i <= sheet.PhysicalNumberOfRows; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null || row.Cells == null || row.Cells.All(a => string.IsNullOrEmpty(a.ToString())))
                {
                    continue;
                }
                var columnIndex = 0;
                var dataRow = table.NewRow();
                foreach (var item in table.Columns)
                {
                    var cell = row.Cells.Where(c => c.ColumnIndex == columnIndex).FirstOrDefault();
                    if (cell == null)
                    {
                        dataRow[item.Name] = null;
                    }
                    else
                    {
                        var value = GetCellValue(cell);
                        dataRow[item.Name] = value;
                    }
                    columnIndex++;
                }
                table.Rows.Add(dataRow);
            }
            return table;
        }

        private object GetCellValue(ICell cell)
        {
            if (cell.CellType == CellType.Numeric)
            {
                if (DateUtil.IsCellDateFormatted(cell))
                {
                    return cell.DateCellValue;
                }
                return cell.NumericCellValue;
            }
            else if (cell.CellType == CellType.Formula)
            {
                return cell.CellFormula;
            }
            else if (cell.CellType == CellType.Boolean)
            {
                return cell.BooleanCellValue;
            }
            else if (cell.CellType == CellType.Blank)
            {
                return null;
            }
            else if (cell.CellType == CellType.String)
            {
                return cell.StringCellValue; ;
            }
            else
            {
                return cell.ToString();
            }

        }
        #endregion

        #region Utilities
        class DefaultExcelStyles
        {
            public ICellStyle ColumnStyle { get; }

            public ICellStyle TitleStyle { get; }

            public DefaultExcelStyles(IWorkbook book)
            {
                var font1 = book.CreateFont();
                font1.IsBold = true;
                font1.FontHeight = 200;
                ColumnStyle = book.CreateCellStyle();
                InitStyle(ColumnStyle);
                ColumnStyle.Alignment = HorizontalAlignment.Center;
                ColumnStyle.VerticalAlignment = VerticalAlignment.Center;
                ColumnStyle.SetFont(font1);

                var font2 = book.CreateFont();
                font2.IsBold = true;
                font2.FontHeight = 350;
                TitleStyle = book.CreateCellStyle();
                InitStyle(TitleStyle);
                TitleStyle.Alignment = HorizontalAlignment.Center;
                TitleStyle.VerticalAlignment = VerticalAlignment.Center;
                TitleStyle.SetFont(font2);
            }

            public static void InitStyle(ICellStyle style)
            {
                style.BorderBottom = BorderStyle.Thin;
                style.BorderTop = BorderStyle.Thin;
                style.BorderLeft = BorderStyle.Thin;
                style.BorderRight = BorderStyle.Thin;
                style.WrapText = true;
            }
        }
        #endregion
    }
}
