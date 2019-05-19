using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTextToDxf
{
    class App
    {
       
        public static void Main(string[] args)
        {
            ExcelReader excel = new ExcelReader();
            excel.getExcelFile();

            DxfWriter dxf = new DxfWriter();
            dxf.dxfWriter();
        }
    }
}
