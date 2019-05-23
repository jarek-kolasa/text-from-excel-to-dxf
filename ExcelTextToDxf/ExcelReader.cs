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
        string[,] textValues;

        public void getExcelFile()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\jkola\Desktop\Programowanie\C#\ExcelTextToCad\test.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

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
                        Console.Write("\r\n"+".");
                    }

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        textValues[i-1, j-1] = xlRange.Cells[i, j].Value2.ToString();
                        // Console.Write(textValues[i-1, j-1]);
                    }
                }
            }


            // Console.ReadLine();


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
            getExcelFile();

            return textValues[row, col];
        }
    }
}
