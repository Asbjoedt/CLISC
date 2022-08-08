﻿using System.IO;
using System.Collections.Generic;
using System.Linq;
//using npoi.xssf.usermodel;
//using npoi.ss.usermodel;
//using npoi.ss.util;
//using npoi.hssf.usermodel;
using Excel = Microsoft.Office.Interop.Excel; 

namespace CLISC
{
    public partial class Conversion
    {

        // Convert legacy Excel files to .xlsx Transitional using Microsoft Office Interop Excel
        public bool Convert_Legacy_ExcelInterop(string input_filepath, string output_filepath)
        {
            bool convert_success = false;

            Excel.Application app = new Excel.Application(); // Create Excel object instance
            app.DisplayAlerts = false; // Don't display any Excel prompts
            Excel.Workbook wb = app.Workbooks.Open(input_filepath, Password:""); // Create workbook instance

            wb.SaveAs(output_filepath, 51); // Save workbook as .xlsx Transitional
            wb.Close(); // Close workbook
            app.Quit(); // Quit Excel application

            convert_success = true; // Mark as succesful
            return convert_success; // Report success
        }

        // Convert to .xlsx Transitional using NPOI
        //public bool Convert_LegacyExcel_NPOI(string org_filepath, string input_filepath, string output_filepath)
        //{
        //    // Read the bytes of the original spreadsheet
        //    FileStream streamed_spreadsheet = new FileStream(input_filepath, FileMode.Open, FileAccess.Read, FileShare.None);
            
        //    // Write the bytes of the converted spreadsheet
        //    byte[] received_spreadsheet = NPOI(streamed_spreadsheet);
        //    MemoryStream spreadsheet = new MemoryStream(received_spreadsheet);
        //    FileStream save_spreadsheet = new FileStream(output_filepath, FileMode.Create, FileAccess.Write);
        //    spreadsheet.WriteTo(save_spreadsheet);
        //    spreadsheet.Close();
        //    save_spreadsheet.Close();

        //    bool convert_success = true;
        //    return convert_success;
        //}

        //// All the methods below are copied from
        //// https://stackoverflow.com/questions/29522173/how-to-save-xls-file-as-xlsx-file-using-npoi-c

        //public static byte[] NPOI(Stream streamed_spreadsheet)
        //{
        //    var source = new HSSFWorkbook(streamed_spreadsheet);
        //    var destination = new XSSFWorkbook();
        //    for (int i = 0; i < source.NumberOfSheets; i++)
        //    {
        //        var xssfSheet = (XSSFSheet)destination.CreateSheet(source.GetSheetAt(i).SheetName);
        //        var hssfSheet = (HSSFSheet)source.GetSheetAt(i);
        //        CopyStyles(hssfSheet, xssfSheet);
        //        CopySheets(hssfSheet, xssfSheet);
        //    }
        //    using (var ms = new MemoryStream())
        //    {
        //        destination.Write(ms);
        //        return ms.ToArray();
        //    }
        //}

        //private static void CopyStyles(HSSFSheet from, XSSFSheet to)
        //{
        //    for (short i = 0; i <= from.Workbook.NumberOfFonts; i++)
        //    {
        //        CopyFont(to.Workbook.CreateFont(), from.Workbook.GetFontAt(i));
        //    }

        //    for (short i = 0; i < from.Workbook.NumCellStyles; i++)
        //    {
        //        CopyStyle(to.Workbook.CreateCellStyle(), from.Workbook.GetCellStyleAt(i), to.Workbook, from.Workbook);
        //    }

        //}

        //private static void CopyFont(IFont toFront, IFont fontFrom)
        //{
        //    toFront.IsBold = fontFrom.IsBold;
        //    toFront.Charset = fontFrom.Charset;
        //    toFront.Color = fontFrom.Color;
        //    toFront.FontHeightInPoints = fontFrom.FontHeightInPoints;
        //    toFront.FontName = fontFrom.FontName;
        //    toFront.IsBold = fontFrom.IsBold;
        //    toFront.IsItalic = fontFrom.IsItalic;
        //    toFront.IsStrikeout = fontFrom.IsStrikeout;
        //    //toFront.Underline = fontFrom.Underline; <- bug in npoi setter
        //}

        //private static void CopyStyle(ICellStyle toCellStyle, ICellStyle fromCellStyle, IWorkbook toWorkbook, IWorkbook fromWorkbook)
        //{
        //    toCellStyle.Alignment = fromCellStyle.Alignment;
        //    toCellStyle.BorderBottom = fromCellStyle.BorderBottom;
        //    toCellStyle.BorderDiagonal = fromCellStyle.BorderDiagonal;
        //    toCellStyle.BorderDiagonalColor = fromCellStyle.BorderDiagonalColor;
        //    toCellStyle.BorderDiagonalLineStyle = fromCellStyle.BorderDiagonalLineStyle;
        //    toCellStyle.BorderLeft = fromCellStyle.BorderLeft;
        //    toCellStyle.BorderRight = fromCellStyle.BorderRight;
        //    toCellStyle.BorderTop = fromCellStyle.BorderTop;
        //    toCellStyle.BottomBorderColor = fromCellStyle.BottomBorderColor;
        //    toCellStyle.DataFormat = fromCellStyle.DataFormat;
        //    toCellStyle.FillBackgroundColor = fromCellStyle.FillBackgroundColor;
        //    toCellStyle.FillForegroundColor = fromCellStyle.FillForegroundColor;
        //    toCellStyle.FillPattern = fromCellStyle.FillPattern;
        //    toCellStyle.Indention = fromCellStyle.Indention;
        //    toCellStyle.IsHidden = fromCellStyle.IsHidden;
        //    toCellStyle.IsLocked = fromCellStyle.IsLocked;
        //    toCellStyle.LeftBorderColor = fromCellStyle.LeftBorderColor;
        //    toCellStyle.RightBorderColor = fromCellStyle.RightBorderColor;
        //    toCellStyle.Rotation = fromCellStyle.Rotation;
        //    toCellStyle.ShrinkToFit = fromCellStyle.ShrinkToFit;
        //    toCellStyle.TopBorderColor = fromCellStyle.TopBorderColor;
        //    toCellStyle.VerticalAlignment = fromCellStyle.VerticalAlignment;
        //    toCellStyle.WrapText = fromCellStyle.WrapText;

        //    toCellStyle.SetFont(toWorkbook.GetFontAt((short)(fromCellStyle.GetFont(fromWorkbook).Index + 1)));
        //}

        //private static void CopySheets(HSSFSheet source, XSSFSheet destination)
        //{
        //    var maxColumnNum = 0;
        //    var mergedRegions = new List<CellRangeAddress>();
        //    var styleMap = new Dictionary<int, HSSFCellStyle>();
        //    for (int i = source.FirstRowNum; i <= source.LastRowNum; i++)
        //    {
        //        var srcRow = (HSSFRow)source.GetRow(i);
        //        var destRow = (XSSFRow)destination.CreateRow(i);
        //        if (srcRow != null)
        //        {
        //            CopyRow(source, destination, srcRow, destRow, mergedRegions);
        //            if (srcRow.LastCellNum > maxColumnNum)
        //            {
        //                maxColumnNum = srcRow.LastCellNum;
        //            }
        //        }
        //    }
        //    for (int i = 0; i <= maxColumnNum; i++)
        //    {
        //        destination.SetColumnWidth(i, source.GetColumnWidth(i));
        //    }
        //}

        //private static void CopyRow(HSSFSheet srcSheet, XSSFSheet destSheet, HSSFRow srcRow, XSSFRow destRow, List<CellRangeAddress> mergedRegions)
        //{
        //    destRow.Height = srcRow.Height;

        //    for (int j = srcRow.FirstCellNum; srcRow.LastCellNum >= 0 && j <= srcRow.LastCellNum; j++)
        //    {
        //        var oldCell = (HSSFCell)srcRow.GetCell(j);
        //        var newCell = (XSSFCell)destRow.GetCell(j);
        //        if (oldCell != null)
        //        {
        //            if (newCell == null)
        //            {
        //                newCell = (XSSFCell)destRow.CreateCell(j);
        //            }

        //            CopyCell(oldCell, newCell);

        //            var mergedRegion = GetMergedRegion(srcSheet, srcRow.RowNum,
        //                    (short)oldCell.ColumnIndex);

        //            if (mergedRegion != null)
        //            {
        //                var newMergedRegion = new CellRangeAddress(mergedRegion.FirstRow,
        //                        mergedRegion.LastRow, mergedRegion.FirstColumn, mergedRegion.LastColumn);
        //                if (IsNewMergedRegion(newMergedRegion, mergedRegions))
        //                {
        //                    mergedRegions.Add(newMergedRegion);
        //                    destSheet.AddMergedRegion(newMergedRegion);
        //                }
        //            }
        //        }
        //    }
        //}

        //private static void CopyCell(HSSFCell oldCell, XSSFCell newCell)
        //{
        //    CopyCellStyle(oldCell, newCell);

        //    CopyCellValue(oldCell, newCell);
        //}

        //private static void CopyCellValue(HSSFCell oldCell, XSSFCell newCell)
        //{
        //    switch (oldCell.CellType)
        //    {
        //        case CellType.String:
        //            newCell.SetCellValue(oldCell.StringCellValue);
        //            break;
        //        case CellType.Numeric:
        //            newCell.SetCellValue(oldCell.NumericCellValue);
        //            break;
        //        case CellType.Blank:
        //            newCell.SetCellType(CellType.Blank);
        //            break;
        //        case CellType.Boolean:
        //            newCell.SetCellValue(oldCell.BooleanCellValue);
        //            break;
        //        case CellType.Error:
        //            newCell.SetCellErrorValue(oldCell.ErrorCellValue);
        //            break;
        //        case CellType.Formula:
        //            newCell.SetCellFormula(oldCell.CellFormula);
        //            break;
        //        default:
        //            break;
        //    }
        //}

        //private static void CopyCellStyle(HSSFCell oldCell, XSSFCell newCell)
        //{
        //    if (oldCell.CellStyle == null) return;
        //    newCell.CellStyle = newCell.Sheet.Workbook.GetCellStyleAt((short)(oldCell.CellStyle.Index + 1));


        //}

        //private static CellRangeAddress GetMergedRegion(HSSFSheet sheet, int rowNum, short cellNum)
        //{
        //    for (var i = 0; i < sheet.NumMergedRegions; i++)
        //    {
        //        var merged = sheet.GetMergedRegion(i);
        //        if (merged.IsInRange(rowNum, cellNum))
        //        {
        //            return merged;
        //        }
        //    }
        //    return null;
        //}

        //private static bool IsNewMergedRegion(CellRangeAddress newMergedRegion,
        //        List<CellRangeAddress> mergedRegions)
        //{
        //    return !mergedRegions.Any(r =>
        //    r.FirstColumn == newMergedRegion.FirstColumn &&
        //    r.LastColumn == newMergedRegion.LastColumn &&
        //    r.FirstRow == newMergedRegion.FirstRow &&
        //    r.LastRow == newMergedRegion.LastRow);
        //}
    }
}