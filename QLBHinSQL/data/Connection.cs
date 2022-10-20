using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace QLBHinSQL.data
{

    internal class Connection
    {
        private SqlConnection sqlConnection;
        private SqlDataAdapter dataAdapter;
        private SqlCommand SqlCommand;
        private string Name_Data = "Data Source=LAPTOP-K9PN94BS\\SQLEXPRESS;Initial Catalog=QLBHCShafp;Integrated Security=True";

        
        public DataTable table(string query)
        {
            DataTable dt = new DataTable();
            using(SqlConnection conn = new SqlConnection(Name_Data))
            {
                conn.Open();
                dataAdapter = new SqlDataAdapter(query, conn);
                dataAdapter.Fill(dt);
                conn.Close();
            }
            return dt;
        }
        public void Excute(string query)
        {
            using(SqlConnection conn = new SqlConnection(Name_Data))
            {
                conn.Open();
                SqlCommand = new SqlCommand(query, conn);
                SqlCommand.ExecuteNonQuery();
                conn.Close();
            }
        }
    }
}
