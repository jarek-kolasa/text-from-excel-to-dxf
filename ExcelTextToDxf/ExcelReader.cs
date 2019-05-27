using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTextToDxf
{
    class ExcelReader
    {
        private string path = @"C:\Users\jkola\Desktop\Programowanie\C#\ExcelTextToCad\test.xlsx";

        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;
        private Excel._Worksheet xlWorksheet;
        private Excel.Range xlRange;

        private int rowCount;
        private int colCount;

        private string[,] textValues;

        protected void GetExcelFile()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(path);
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;

            rowCount = xlRange.Rows.Count;
            colCount = xlRange.Columns.Count;

             textValues = new string[rowCount, colCount];

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    // new line
                    if (j == 1)
                    {
                        Console.Write("."+ "\r\n");
                    }

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        textValues[i-1, j-1] = xlRange.Cells[i, j].Value2.ToString();
                        // Console.Write(textValues[i-1, j-1]);
                    }
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            // close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            // quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }


        public string GetChoosenCellValue(int row, int col)
        {
            GetExcelFile();

            return textValues[row, col];
        }
    }
}
