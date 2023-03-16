
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CADVISER_Transcript_Customers.Model.Data_Set.Index_Set;
using CADVISER_Transcript_Customers.Model;
namespace CADVISER_Transcript_Customers.Controller
{
    public class Controller_SQL_Function
    {
        Function_SQL function_sql = new Function_SQL();


        public List<string> Controller_SQL_Action_Function(Index_SQL_Action_Category mode_index, Index_SQL_Action_Function query_index, SqlConnection conn, string value)
        {

            List<string> Result = new List<string>();

            try
            {
                switch (query_index)
                {
                    case Index_SQL_Action_Function.Try_Connection:
                        return function_sql.Connection(conn,value);

                    case Index_SQL_Action_Function.Get_Building:
                        return function_sql.Get_Building(mode_index, query_index, conn, value);

                    case Index_SQL_Action_Function.Get_Owner:
                        return function_sql.Get_Owner(mode_index, query_index, conn, value);

                    case Index_SQL_Action_Function.Get_Loan_Detail:
                        return function_sql.Get_Loan_Details(mode_index,query_index, conn, value);

                    case Index_SQL_Action_Function.Get_Building_Type:
                        return function_sql.Get_Building_Base_Data(mode_index, query_index, conn, value);

                    default:
                        break;
                }



            }
            catch (Exception)
            {

                Result.Add("發生異常問題，請聯絡【研發中心-郁宸】。");
            }




            return Result;
        }



    }
}
