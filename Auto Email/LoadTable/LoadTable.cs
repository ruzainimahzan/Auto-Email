using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBClass
{
    public class LoadTable
    {
        static string dbcon = ConfigurationManager.ConnectionStrings["dbAutoEmail"].ConnectionString;

        public static bool FetchEmail(string FEmailID)
        {
            bool LPass = false;
            string sqlSelect = ConfigurationManager.AppSettings["SQLViewEmail"];
            SqlDataAdapter da = new SqlDataAdapter(sqlSelect, dbcon);
            da.TableMappings.Add("Table", "Email_Detail");

            // Create and fill the DataSet
            DataSet ds = new DataSet();
            da.Fill(ds);

            DataTable EmailInfo = ds.Tables["Email_Detail"];
            var EmailQuery = from d in EmailInfo.AsEnumerable()
                              where d.Field<string>("email_ID") == FEmailID
                              select d;

            foreach (var e in EmailQuery)
            {
                LPass = true;
            }

            return LPass;
        }

    }

    public 
}
