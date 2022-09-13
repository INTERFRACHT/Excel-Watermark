using System;
using System.IO;

namespace Excel_Watermark
{
    class FileHandler
    {
        private ErrorHandler ErrorHandler = new ErrorHandler();

        public void ProcessFiles(string sourcePath, string destinationPath = "")
        {
            ConvertFileToPdf(sourcePath, destinationPath);
        }
        
        /// <summary>
        /// Spracuje subor
        /// </summary>
        /// <param name="fileName"></param>
        private void ConvertFileToPdf(string sourcePath, string destinationPath = "")
        {
            try
            {
                string fileName = Path.GetFileName(sourcePath);
                string outputFilePath = "";

                if (destinationPath == "")
                    outputFilePath = sourcePath.Replace(".xlsx", ".pdf");
                else
                    outputFilePath = destinationPath.Replace(".xlsx", ".pdf");

                // Create COM Objects
                Microsoft.Office.Interop.Excel.Application excelApplication;
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;

                // Create new instance of Excel
                excelApplication = new Microsoft.Office.Interop.Excel.Application();

                // Make the process invisible to the user
                excelApplication.ScreenUpdating = false;

                // Make the process silent
                excelApplication.DisplayAlerts = false;

                // Open the workbook that you wish to export to PDF
                excelWorkbook = excelApplication.Workbooks.Open(sourcePath);

                // If the workbook failed to open, stop, clean up, and bail out
                if (excelWorkbook == null)
                {
                    excelApplication.Quit();

                    excelApplication = null;
                    excelWorkbook = null;
                }

                try
                {
                    // Call Excel's native export function (valid in Office 2007 and Office 2010, AFAIK)
                    excelWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputFilePath);
                }
                catch (System.Exception ex)
                {
                    //TODO show error     
                }
                finally
                {
                    // Close the workbook, quit the Excel, and clean up regardless of the results...
                    excelWorkbook.Close();
                    excelApplication.Quit();
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.SendError("ConvertFileToPdf", ex.ToString());
                Console.WriteLine(ex.ToString());
            }
        }
    }
}