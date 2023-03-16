using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace CADVISER_Transcript_Customers.Model.Data_Set
{
    public class Set_SQL_Conn
    {
        public string Data_Source { get; set; }
        public string Data_Base { get; set; }
        public string User_id { get; set; }
        public string Password { get; set; }
        public string conn_str { get; set; }

        public void Load_SQLConn_Str_CADVVISER_DB()
        {
            
            Data_Source = "202.39.78.130";
            Data_Base = "全城地產-city-tech.tw";
            User_id = "PTMB-YC";
            Password = "Qib0808";
            conn_str = string.Format("Data Source={0}; Database={1}; Trusted_Connection={2};user id={3};password={4}", Data_Source, Data_Base, "false", User_id, Password);
        }
        public void Load_SQLConn_Str_Temp_DB()
        {

            Data_Source = "localhost";
            Data_Base = "CADVISER_";
            User_id = "PTMB-YC";
            Password = "Qib0808";
            conn_str = string.Format("Data Source={0}; Database={1}; Trusted_Connection={2};user id={3};password={4}", Data_Source, Data_Base, "false", User_id, Password);
        }

        public SqlConnection Get_Conn(string conn_str)
        {
            return new SqlConnection(conn_str);
        }
    }
}
