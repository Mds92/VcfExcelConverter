using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace VcfConverter.Classes
{
    public static class ExcelHelper
    {
        private static readonly Font Font = new("Segoe UI", 10);
        private static readonly Color BorderColor = Color.FromArgb(161, 161, 161);
        private static void FormatExcelRange(ExcelRange excelRange, bool setBackgroundColor = false)
        {
            excelRange.Style.Font.SetFromFont(Font);
            excelRange.Style.Border.BorderAround(ExcelBorderStyle.Thin, BorderColor);
            excelRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelRange.Style.Fill.BackgroundColor.SetColor(Color.White);
            excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelRange.Style.ShrinkToFit = true;
            excelRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            excelRange.Style.WrapText = true;
            if (setBackgroundColor)
                excelRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(237, 237, 237));
        }
        private static string FromIntegerIndexToColumnLetter(int intCol)
        {
            var intFirstLetter = intCol / 676 + 64;
            var intSecondLetter = intCol % 676 / 26 + 64;
            var intThirdLetter = intCol % 26 + 65;

            var firstLetter = intFirstLetter > 64 ? (char)intFirstLetter : ' ';
            var secondLetter = intSecondLetter > 64 ? (char)intSecondLetter : ' ';
            var thirdLetter = (char)intThirdLetter;

            return string.Concat(firstLetter, secondLetter, thirdLetter).Trim();
        }

        public static byte[] GetExcelFile(string sheetTitle, bool rightToLeft, ExcelFileModel model)
        {
            using var excelPackage = new ExcelPackage();
            var rowNumber = 1;
            var colNumber = 0;
            var excelWorksheet = excelPackage.Workbook.Worksheets.Add(sheetTitle);
            excelWorksheet.View.RightToLeft = rightToLeft;
            excelWorksheet.Row(1).Height = 25;
            foreach (var cellModel in model.Cells.Where(q => q != null))
            {
                colNumber++;
                var cellAddress = $"{FromIntegerIndexToColumnLetter(colNumber - 1)}{rowNumber}";
                excelWorksheet.Cells[cellAddress].Value = cellModel.Title;
                FormatExcelRange(excelWorksheet.Cells[cellAddress], true);
                excelWorksheet.Cells[cellAddress].Style.Font.Bold = true;
                if (cellModel.Width > 0)
                    excelWorksheet.Column(colNumber).Width = cellModel.Width;
                else
                    excelWorksheet.Column(colNumber).AutoFit();
            }
            colNumber = 1;
            foreach (var rowCells in model.Rows)
            {
                rowNumber++;
                foreach (var cell in rowCells)
                {
                    if (model.Cells.Count < colNumber) continue;
                    var cellModel = model.Cells[colNumber - 1];
                    var cellAddress = $"{FromIntegerIndexToColumnLetter(colNumber - 1)}{rowNumber}";
                    excelWorksheet.Cells[cellAddress].Value = cell;
                    FormatExcelRange(excelWorksheet.Cells[cellAddress]);
                    if (!string.IsNullOrWhiteSpace(cellModel?.NumberFormat))
                        excelWorksheet.Cells[cellAddress].Style.Numberformat.Format = cellModel.NumberFormat;
                    colNumber++;
                }
                colNumber = 1;
            }
            return excelPackage.GetAsByteArray();
        }
        public static List<Dictionary<string, dynamic>> GetExcelRowsData(string excelFilePath)
        {
            var rowsData = new List<Dictionary<string, dynamic>>();
            var fi = new FileInfo(excelFilePath);
            using var excelPackage = new ExcelPackage(fi);
            var excelColumnList = new List<KeyValuePair<string, string>>();
            var excelWorksheet = excelPackage.Workbook.Worksheets[0];
            foreach (var cell in excelWorksheet.Cells[excelWorksheet.Dimension.Start.Row, excelWorksheet.Dimension.Start.Column, 1, excelWorksheet.Dimension.End.Column])
                excelColumnList.Add(new KeyValuePair<string, string>(cell.Text, Regex.Replace(cell.Address, @"\d+", "")));
            for (var i = 2; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                var rowData = new Dictionary<string, dynamic>();

                // ایجاد دیکشنری از یک ردیف اطلاعات موجود در اکسل
                for (var j = 1; j <= excelWorksheet.Dimension.End.Column; j++)
                {
                    var cell = excelWorksheet.Cells[i, j];
                    var cellValue = excelWorksheet.Cells[i, j].Value;
                    if (cell.Address.IndexOf(":", StringComparison.InvariantCultureIgnoreCase) > -1)
                        throw new Exception("فایل اکسل نباید دارای ستون های ادغام شده باشد");
                    var columnName = Regex.Replace(cell.Address, @"\d+", "");
                    var (propertyName, _) = excelColumnList.FirstOrDefault(q => q.Value.Equals(columnName, StringComparison.InvariantCultureIgnoreCase));
                    if (string.IsNullOrWhiteSpace(propertyName)) continue;
                    rowData.Add(propertyName, cellValue);
                }
                rowsData.Add(rowData);
            }
            return rowsData;
        }
    }
}
