using Microsoft.Win32.SafeHandles;
using System;
using System.IO;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;
using System.Threading;

namespace Excel_Watermark
{
    class FileHandler
    {
        private ErrorHandler ErrorHandler = new ErrorHandler();

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool LogonUser(String lpszUsername, String lpszDomain, String lpszPassword,
        int dwLogonType, int dwLogonProvider, out SafeTokenHandle phToken);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public extern static bool CloseHandle(IntPtr handle);


        public int ProcessFiles(string sourcePath, string destinationPath = "")
        {
            return ConvertFileToPdf(sourcePath, destinationPath);
        }

        public sealed class SafeTokenHandle : SafeHandleZeroOrMinusOneIsInvalid
        {
            private SafeTokenHandle()
                : base(true)
            {
            }

            [DllImport("kernel32.dll")]
            [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
            [SuppressUnmanagedCodeSecurity]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool CloseHandle(IntPtr handle);

            protected override bool ReleaseHandle()
            {
                return CloseHandle(handle);
            }
        }

        /// <summary>
        /// Impersonacia kvoli pristupu
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="destinationPath"></param>
        private void Impersonate(string sourcePath, string destinationPath)
        {
            SafeTokenHandle safeTokenHandle;
            try
            {
                string domainName = "INTERFRACHT";
                string userName = "AdminIT1";
                string password = "Republika 1";

                const int LOGON32_PROVIDER_DEFAULT = 0;
                const int LOGON32_LOGON_INTERACTIVE = 2;

                // Call LogonUser to obtain a handle to an access token.
                bool returnValue = LogonUser(userName, domainName, password,
                    LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT,
                    out safeTokenHandle);

                if (false == returnValue)
                {
                    int ret = Marshal.GetLastWin32Error();
                    throw new System.ComponentModel.Win32Exception(ret);
                }
                using (safeTokenHandle)
                {
                    using (WindowsIdentity newId = new WindowsIdentity(safeTokenHandle.DangerousGetHandle()))
                    {
                        using (WindowsImpersonationContext impersonatedUser = newId.Impersonate())
                        {
                            ConvertFileToPdf(sourcePath, destinationPath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.SendError("Impersonate", ex.ToString());
            }
        }

        /// <summary>
        /// Spracuje subor
        /// </summary>
        /// <param name="fileName"></param>
        private int ConvertFileToPdf(string sourcePath, string destinationPath = "")
        {
            int result = 0;

            try
            {
                string fileName = Path.GetFileName(sourcePath);
                string outputFilePath = "";

                if (destinationPath == "")
                    outputFilePath = sourcePath.Replace(".xlsx", ".pdf");
                else
                    outputFilePath = destinationPath.Replace(".xlsx", ".pdf");

                //vymazanie suboru ak existuje
                if (File.Exists(outputFilePath))
                {
                    File.Delete(outputFilePath);
                    result++;
                    Thread.Sleep(5000);
                }

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

            return result;
        }
    }
}