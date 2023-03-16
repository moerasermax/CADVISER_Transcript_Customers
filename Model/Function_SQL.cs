using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using CADVISER_Transcript_Customers.Model.Data_Set;
using CADVISER_Transcript_Customers.Model.Data_Set.Index_Set;
namespace CADVISER_Transcript_Customers.Model
{
    public class Function_SQL
    {
        Operation_SQL operation_sql = new Operation_SQL();
        public List<string> Connection(SqlConnection conn,string DB_Base )
        {
            List<string> result = new List<string>();

            try
            {
                conn.Open();
                conn.Close();

                result.Add(string.Format("已成功連線至【{0}】伺服器\r\n", DB_Base));
            }
            catch (Exception)
            {

                result.Add(string.Format("連線失敗，請聯絡【研發中心-郁宸】\r\n"));
            }


            return result;
        }

        public List<string> Get_Building(Index_SQL_Action_Category mode_index, Index_SQL_Action_Function query_index, SqlConnection conn, string value)
        {
            List<string> result = new List<string>();


                string cmd_str = string.Format("SELECT [BuildingBaseId],[建物門牌],[層數],[總面積] FROM [全城地產-city-tech.tw].[dbo].[Building_Building] where 建物坐落地號 != '無提供資料' and 建物坐落地號 != '' order by UpdateTime desc");
                SqlCommand cmd = new SqlCommand(cmd_str, conn);
                result = operation_sql.SQL_Excute(mode_index, query_index, conn, cmd);

            return result;
        }

        public List<string> Get_Owner(Index_SQL_Action_Category mode_index, Index_SQL_Action_Function query_index, SqlConnection conn, string value)
        {
            List<string> result = new List<string>();

                string cmd_str = string.Format("SELECT [所有權人姓名],[統一編號],[登記原因],[權利範圍],[相關他項登記次序] FROM [全城地產-city-tech.tw].[dbo].[Building_BuildingOwner] where BuildingBaseId=@BuildingBaseId And [相關他項登記次序] != ''");
                SqlCommand cmd = new SqlCommand(cmd_str, conn);
                operation_sql.SQL_PARAMETER_Insert(query_index, cmd, value);
                result = operation_sql.SQL_Excute(mode_index, query_index, conn, cmd);

            return result;
        }

        public List<string> Get_Loan_Details(Index_SQL_Action_Category mode_index, Index_SQL_Action_Function query_index, SqlConnection conn, string value)
        {
            List<string> result = new List<string>();

                string cmd_str = string.Format("SELECT [登記日期] ,[權利人姓名] ,[擔保債權總金額],[他項登記次序] FROM [全城地產-city-tech.tw].[dbo].[Building_BuildingOther] where BuildingBaseId = @BuildingBaseId and 他項登記次序 = @他項登記次序");
                SqlCommand cmd = new SqlCommand(cmd_str, conn);
                operation_sql.SQL_PARAMETER_Insert(query_index, cmd, value);
                result = operation_sql.SQL_Excute(mode_index, query_index, conn, cmd);


            return result;
        }

        public List<string> Get_Building_Base_Data(Index_SQL_Action_Category mode_index, Index_SQL_Action_Function query_index, SqlConnection conn, string value)
        {
            List<string> result = new List<string>();

                string cmd_str = string.Format("SELECT [City],[Township],[建物類型],[Section],[BuildingNumber]  FROM [全城地產-city-tech.tw].[dbo].[Building_BuildingBase]  where ID = @ID");
                SqlCommand cmd = new SqlCommand(cmd_str, conn);
                operation_sql.SQL_PARAMETER_Insert(query_index, cmd, value);
                result = operation_sql.SQL_Excute(mode_index, query_index, conn, cmd);

            return result;
        }
    }

    




    public class Operation_SQL
    {
        public List<string> SQL_Excute(Index_SQL_Action_Category mode_index, Index_SQL_Action_Function query_index, SqlConnection conn, SqlCommand cmd)
        {
            List<string> result = new List<string>();
            Function_DataProcess data_process = new Function_DataProcess();

            try
            {
                switch (mode_index)
                {
                    case Index_SQL_Action_Category.Get:
                        DataTable dataTable = new DataTable();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dataTable);
                        conn.Close();
                        da.Dispose();
                        result = data_process.Process_SQL_Result(query_index, dataTable);
                        break;
                   

                    default:
                        result.Add("搜尋不到指定功能，請聯絡【研發中心-郁宸】。");
                        break;
                }
            }
            catch (Exception e)
            {
                result.Add("執行查詢時發生問題，請聯絡【研發中心-郁宸】。");
                MessageBox.Show( e.Message.ToString());
            }

            return result;


        }

        public SqlCommand SQL_PARAMETER_Insert(Index_SQL_Action_Function query_index, SqlCommand cmd, string value)
        {
            string[] value_arr;
            switch (query_index)
            {
                case Index_SQL_Action_Function.Get_Owner:
                    cmd.Parameters.AddWithValue("@BuildingBaseId", value);
                    return cmd;
                case Index_SQL_Action_Function.Get_Loan_Detail:
                    value_arr = value.Split(',');
                    cmd.Parameters.AddWithValue("@BuildingBaseId", value_arr[0]);
                    cmd.Parameters.AddWithValue("@他項登記次序", value_arr[1]);
                    return cmd;
                case Index_SQL_Action_Function.Get_Building_Type:
                    cmd.Parameters.AddWithValue("@ID", value);
                    return cmd; 
                default:
                    return cmd;
            }



        }
    }
}
