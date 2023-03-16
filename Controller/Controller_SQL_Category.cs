using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CADVISER_Transcript_Customers.Model.Data_Set.Index_Set;

namespace CADVISER_Transcript_Customers.Controller
{
    public class Controller_SQL_Category
    {

        Controller_SQL_Function sql_function_controller = new Controller_SQL_Function();

        public List<string> Controller_SQL_Action_Category(Index_SQL_Action_Category mode_index,Index_SQL_Action_Function query_index, SqlConnection conn, string value)
        {
            List<string> Result = new List<string>();

            try
            {
                switch (mode_index)
                {
                    case Index_SQL_Action_Category.Test:
                        return sql_function_controller.Controller_SQL_Action_Function(mode_index,query_index, conn,value);
                    case Index_SQL_Action_Category.Get:
                        return sql_function_controller.Controller_SQL_Action_Function(mode_index,query_index, conn, value);
                    case Index_SQL_Action_Category.Insert:
                        return sql_function_controller.Controller_SQL_Action_Function(mode_index, query_index, conn, value);
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
