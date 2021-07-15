using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
/*
利用Oledb操作Access mdb檔案
https://einboch.pixnet.net/blog/post/245703881

*/
using System.Data;
using System.Data.OleDb;

namespace CS_Console_AccessDB
{
    class Program
    {
        public static OleDbConnection OleDbOpenConn(string Database= "CS_MDB_test.MDB")
        {
            string cnstr = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Database);
            OleDbConnection icn = new OleDbConnection();
            icn.ConnectionString = cnstr;
            if (icn.State == ConnectionState.Open) icn.Close();
            icn.Open();
            return icn;
        }
        public static DataTable GetOleDbDataTable(string Database, string OleDbString)
        {
            DataTable myDataTable = new DataTable();
            OleDbConnection icn = OleDbOpenConn(Database);
            OleDbDataAdapter da = new OleDbDataAdapter(OleDbString, icn);
            DataSet ds = new DataSet();
            ds.Clear();
            da.Fill(ds);
            myDataTable = ds.Tables[0];
            if (icn.State == ConnectionState.Open) icn.Close();
            return myDataTable;
        }
        public static void OleDbInsertUpdateDelete(string Database, string OleDbSelectString)
        {
            string cnstr = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Database);
            OleDbConnection icn = OleDbOpenConn(cnstr);
            OleDbCommand cmd = new OleDbCommand(OleDbSelectString, icn);
            cmd.ExecuteNonQuery();
            if (icn.State == ConnectionState.Open) icn.Close();
        }
        static void Pause()
        {
            Console.Write("Press any key to continue...");
            Console.ReadKey(true);
        }
        static void Main(string[] args)
        {
            string sql = "DELETE FROM Table01;";
            OleDbInsertUpdateDelete("CS_MDB_test.MDB", sql);
            sql = "INSERT INTO Table01 (id, name) VALUES(1,'jash');";
            OleDbInsertUpdateDelete("CS_MDB_test.MDB", sql);
            sql = "INSERT INTO Table01 (id, name) VALUES(2,'jash.liao');";
            OleDbInsertUpdateDelete("CS_MDB_test.MDB", sql);
            sql = "INSERT INTO Table01 (id, name) VALUES(3,'liao.jash');";
            OleDbInsertUpdateDelete("CS_MDB_test.MDB", sql);

            sql = "SELECT * FROM Table01";//3 rows
            sql = "SELECT * FROM Table01 WHERE name LIKE 'jash%'";//2 rows
            DataTable dt = GetOleDbDataTable("CS_MDB_test.MDB", sql);
            for(int i=0;i< dt.Rows.Count;i++)
            {
                Console.WriteLine("{0},{1}", dt.Rows[i]["id"].ToString(), dt.Rows[i]["name"].ToString());
            }

            dt.Clear();
            Pause();
        }
    }
}
