using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KaizhongRPA
{
    public class SqlClass
    {        
        public int InsSQL(string sql, string connstr)
        {
            int result = -1;
            using (SqlConnection conn = new SqlConnection(connstr))
            {
                try
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    result = cmd.ExecuteNonQuery();
                }
                catch { }
                finally { conn.Close(); }
            }
            return result;
        }

       
        public int DelSQL(string sql, string connstr)
        {
            int result = -1;
            using (SqlConnection conn = new SqlConnection(connstr))
            {
                try
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    result = cmd.ExecuteNonQuery();
                }
                catch { }
                finally { conn.Close(); }
            }
            return result;
        }
     
        public DataTable SlcSQL(string sql, string connstr)
        {
            DataTable result = null;
            using (SqlConnection conn = new SqlConnection(connstr))
            {
                try
                {
                    conn.Open();
                    SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "slc_of_sql");
                    result = ds.Tables["slc_of_sql"];
                }
                catch { }
                finally { conn.Close(); }
            }
            return result;
        }
      
        public int UpdSQL(string sql, string connstr)
        {
            int result = -1;
            using (SqlConnection conn = new SqlConnection(connstr))
            {
                try
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    result = cmd.ExecuteNonQuery();
                }
                catch { }
                finally { conn.Close(); }
            }
            return result;          
        }

    }
}
