namespace CLISCC.Classes
{
    public class Convert_XLS
    {
        // Source: https://lockevn.medium.com/convert-excel-format-xls-to-xlsx-using-c-3e1a348fca22
        // <param name="filesFolder"></param>
        public static string ConvertXLS_XLSX(FileInfo file)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            var xlsFile = file.FullName;
            var wb = app.Workbooks.Open(xlsFile);
            var xlsxFile = xlsFile + "x";
            wb.SaveAs(Filename: xlsxFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();
            return xlsxFile;
        }
    }
}