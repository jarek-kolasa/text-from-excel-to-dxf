using netDxf;
using netDxf.Entities;
using netDxf.Header;
using netDxf.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;


namespace ExcelTextToDxf
{
    class DxfWriter
    {

        // dxf filename
        string [] filesPath = Directory.GetFiles(@"C:\Users\jkola\Desktop\Programowanie\C#\", "*.dxf");

        // by default it will create an AutoCad2000 DXF version
        DxfDocument dxfDocument;

        // an entity
        Line entity;

        // text
        Vector3 textLocation = new Vector3(0, 0, 0);
        Text text;

        public void dxfWriter()
        {
            dxfDocument = new DxfDocument();
            entity = new Line(new Vector2(5, 5), new Vector2(10, 5));
            //add an entity here
            dxfDocument.AddEntity(entity);
            // text
            text = new Text("text", textLocation, 2.0);
            Layer layer = new Layer("text");
            text.Layer = layer;
            text.Alignment = TextAlignment.BottomLeft;
            dxfDocument.AddEntity(text);
            // save to file

            foreach (string file in filesPath)
            {
            dxfDocument.Save(file);
            }

            bool isBinary;
            foreach (string file in filesPath)
            {
            // this check is optional but recommended before loading a DXF file
            DxfVersion dxfVersion = DxfDocument.CheckDxfFileVersion(file, out isBinary);
            // netDxf is only compatible with AutoCad2000 and higher DXF version
            if (dxfVersion < DxfVersion.AutoCad2000) return;
            // load file
            DxfDocument loaded = DxfDocument.Load(file);
            }
        }
    }
}

