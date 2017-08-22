using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_utils
{
    public class DataAccess
    {
        public DataAccess(string connString, string queryType)
        {

        }

        public DataSet GetDataSet(string sQLQuery, object[] parms, string[] v)
        {
            return new DataSet();
        }

        public OleDbType SQLType(Type dataType)
        {
            throw new NotImplementedException();
        }
    }
}
