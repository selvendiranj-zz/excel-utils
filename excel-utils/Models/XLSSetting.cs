using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_utils.Models
{
    public class XLSSetting
    {
        private string connString = "" +
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};" +
            "Extended Properties={2}Excel 8.0;HDR={1}{2}";
        private string fileName = "";
        private string hasHeader = "Yes";
        private string sheets = "Sheet1";
        private string fntName = "Arial";
        private bool isNew = true;
        private bool delRow = false;
        private int fntSize = 10;
        private int zoomPct = 100;
        private int hdrPosn = 1;
        private bool isFormat = false;
        private string format = "";

        public string ConnString { get => connString; set => connString = value; }
        public string FileName { get => fileName; set => fileName = value; }
        public string HasHeader { get => hasHeader; set => hasHeader = value; }
        public string Sheets { get => sheets; set => sheets = value; }
        public string FntName { get => fntName; set => fntName = value; }
        public bool IsNew { get => isNew; set => isNew = value; }
        public bool DelRow { get => delRow; set => delRow = value; }
        public int FntSize { get => fntSize; set => fntSize = value; }
        public int ZoomPct { get => zoomPct; set => zoomPct = value; }
        public int HdrPosn { get => hdrPosn; set => hdrPosn = value; }
        public bool IsFormat { get => isFormat; set => isFormat = value; }
        public string Format { get => format; set => format = value; }
    }
}
