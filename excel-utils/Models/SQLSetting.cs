using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_utils.Models
{
    public class SQLSetting
    {
        private string connString = "" +
            "Data Source={0};Initial Catalog={1};" +
            "Integrated Security=SSPI;Persist Security Info=False;" +
            "Network Library=DBMSSOCN;Packet Size=8192;Max Pool Size=200;";
        private string server = "";
        private string database = "";
        private string query = "";
        private string queryType = "Text";
        private string errFile = "";
        private string parms = "";
        private object[] parmCollection = null;

        public string ConnString { get => connString; set => connString = value; }
        public string Server { get => server; set => server = value; }
        public string Database { get => database; set => database = value; }
        public string Query { get => query; set => query = value; }
        public string QueryType { get => queryType; set => queryType = value; }
        public string ErrFile { get => errFile; set => errFile = value; }
        public string Parms { get => parms; set => parms = value; }
        public object[] ParmCollection { get => parmCollection; set => parmCollection = value; }
    }
}
