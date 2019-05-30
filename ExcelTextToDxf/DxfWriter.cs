using netDxf;
using netDxf.Entities;
using netDxf.Header;
using netDxf.Tables;
using System;
using System.IO;


namespace ExcelTextToDxf
{
    class DxfWriter
    {

        // dxf filename
        private string [] filesPath;

        // by default it will create an AutoCad2000 DXF version
        private DxfDocument dxfDocument;

        // an entity
        private Line entity;

        // text
        private Vector3 textLocation = new Vector3(0, 0, 0);
        private Text text;

        // one object of ExcelReader to read values
        private ExcelReader excelText;

        // TODO - console.read choosen cells
        private int excelRow = 0;
        private int excelCol = 1;

        public void DxfWriterApp()
        {
            // path of dxf file
            Console.Write("Podaj sciezke do plikow *.dxf (np. C:\\Users\\user\\Desktop\\): ");
            filesPath = Directory.GetFiles(Console.ReadLine(), "*.dxf");

            // object of ExcelReader
            excelText = new ExcelReader();
            string readExcelText = excelText.GetChoosenCellValue(excelRow, excelCol);

            bool isBinary;
            foreach (string file in filesPath)
            {
                // this check is optional but recommended before loading a DXF file
                DxfVersion dxfVersion = DxfDocument.CheckDxfFileVersion(file, out isBinary);
                // netDxf is only compatible with AutoCad2000 and higher DXF version
                if (dxfVersion < DxfVersion.AutoCad2000) return;
                // load file
                dxfDocument = DxfDocument.Load(file);

                entity = new Line(new Vector2(5, 5), new Vector2(10, 5));
                //add an entity here
                dxfDocument.AddEntity(entity);
                // text
                text = new Text(readExcelText, textLocation, 2.0);
                Layer layer = new Layer("text");
                text.Layer = layer;
                text.Alignment = TextAlignment.BottomLeft;
                dxfDocument.AddEntity(text);
                // save to file
                dxfDocument.Save(file);
            }

        }
    }
}

