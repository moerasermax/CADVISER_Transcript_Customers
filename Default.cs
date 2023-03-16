using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


using CADVISER_Transcript_Customers.Model;
using CADVISER_Transcript_Customers.Controller;
using CADVISER_Transcript_Customers.Model.Data_Set;
using CADVISER_Transcript_Customers.Model.Data_Set.Index_Set;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Newtonsoft.Json;

namespace CADVISER_Transcript_Customers
{
    public partial class Default : Form
    {
        Function_Load_Data function_load_data = new Function_Load_Data();
        List<Set_Pre_Customer_Data> list_pre_customers = new List<Set_Pre_Customer_Data>();

        public Default()
        {
            InitializeComponent();
        }

        private void Get_Data_Click(object sender, EventArgs e)
        {

            Load_Data();
            Evaluate();
            Export();
            Log.Text += JsonConvert.SerializeObject(list_pre_customers);


        }

        #region 測試區
        private void Try_CADVISER_DB_Click(object sender, EventArgs e)
        {
            Log.Text += function_load_data.Test_CADVVISER();
        }

        private void Try_Temp_DB_Click(object sender, EventArgs e)
        {
            Log.Text += function_load_data.Test_CADVVISER();
        }
        #endregion


        #region 功能區

        private void Load_Data()
        {
            string Log_Status_str = "";

            function_load_data.Set_Load_Data();
            list_pre_customers = function_load_data.Load_All_Data(ref  Log_Status_str);

            Log.Text = Log_Status_str;
        }

        private void Evaluate()
        {
            Function_Evaluate_Data function_evaluate = new Function_Evaluate_Data();
            function_evaluate.Set_pre_customers_data(list_pre_customers);
            function_evaluate.Evaluate_Score();
        }

        private void Export()
        {
            Function_Export_Excel function_export_excel = new Function_Export_Excel();
            function_export_excel.Export_Excel(list_pre_customers);
        }

        #endregion

        private void Test_Data_Click(object sender, EventArgs e)
        {
        }
    }


    class Function_Load_Data
    {
        Controller_SQL_Category controller_sql_categoryt = new Controller_SQL_Category();
        Function_DataProcess function_data_process = new Function_DataProcess();
        Set_SQL_Conn set_sql_conn = new Set_SQL_Conn();
        SqlConnection conn;
        List<Set_Pre_Customer_Data> temp_list_pre_customers = new List<Set_Pre_Customer_Data>();

        List<string> data = new List<string>();


        public void Set_Load_Data()
        {
            set_sql_conn.Load_SQLConn_Str_CADVVISER_DB();
            conn =new SqlConnection(set_sql_conn.conn_str);
        }

        public List<Set_Pre_Customer_Data> Load_All_Data(ref string Log_Status)
        {
            Load_BuildingBaseID();  /// 撈「所有建物」的基本資料。
            Log_Status = string.Format("{0}{1}" , Log_Status, data[0]);
            Load_Owner_Data();  /// 撈「建物的所有權者」資料。
            Log_Status = string.Format("{0}{1}", Log_Status, data[0]);
            Load_Loan_Data();  /// 撈「所有權者的債務」資料。
            Log_Status = string.Format("{0}{1}", Log_Status, data[0]);
            Load_Building_Base_Data();  /// 補全「所有建物」基本資料，補全內容：所在城市、區、建物類型。
            Log_Status = string.Format("{0}{1}", Log_Status, data[0]);
            Load_Updating_Time_Data();  /// 紀載 「現在時間」作為「更新時間」。
            Log_Status = string.Format("{0}{1}", Log_Status, data[0]);
            return temp_list_pre_customers;
        }

        public void Load_BuildingBaseID()
        {
            try
            {
                /// 建築物資訊
                /// 資料內容有：[BuildingBaseId],[登記日期],[登記原因],[層數],[總面積]
                data = controller_sql_categoryt.Controller_SQL_Action_Category(Index_SQL_Action_Category.Get, Index_SQL_Action_Function.Get_Building, conn, "");
                data.Insert(0, string.Format("成功獲取「建築編號」資料。\r\n"));
                function_data_process.Set_BuildingBaseID(ref temp_list_pre_customers, data);
            }
            catch (Exception)
            {

                data.Insert(0, string.Format("獲取「建築編號」資料失敗，請聯絡【研發中心-郁宸】。\r\n"));
            }

        }

        public void Load_Owner_Data()
        {
            try
            {
                /// 擁有者資料
                /// 資料內容有：[客戶姓名],[性別],[身份證號],[登記原因],[權利範圍],[相關他項登記次序]
                for (int i = 0; i <= temp_list_pre_customers.Count - 1; i++)
                {
                    data = controller_sql_categoryt.Controller_SQL_Action_Category(Index_SQL_Action_Category.Get, Index_SQL_Action_Function.Get_Owner, conn, temp_list_pre_customers[i].buildingbaseId);
                    data.Insert(0, string.Format("成功獲取「所有權人」資料。\r\n"));
                    function_data_process.Set_Customer_Data(ref temp_list_pre_customers, data, i);
                }


            }
            catch (Exception)
            {

                data.Insert(0, string.Format("獲取「所有權人」資料失敗，請聯絡【研發中心-郁宸】。\r\n"));

            }

        }

        public void Load_Loan_Data() /// 他項 ---> other
        {
            try
            {
                /// 擁有者資料
                /// 資料內容有：【胎數資訊】
                for (int i = 0; i <= temp_list_pre_customers.Count - 1; i++)
                {
                    for (int j = 0; j <= temp_list_pre_customers[i].customer_data.Count - 1; j++)
                    {
                        data.Clear();
                        string[] order_tittle_arr = temp_list_pre_customers[i].customer_data[j].loan_tittle.Split('｜');
                        for (int x = 0; x <= order_tittle_arr.Length - 1; x++)
                        {
                            List<string> single_loan_detail_data = new List<string>();
                            string value = temp_list_pre_customers[i].buildingbaseId + "," + order_tittle_arr[x];
                            single_loan_detail_data = controller_sql_categoryt.Controller_SQL_Action_Category(Index_SQL_Action_Category.Get, Index_SQL_Action_Function.Get_Loan_Detail, conn, value);
                           
                            ///防止他項裡面有空值狀況
                            if(single_loan_detail_data.Count != 0)
                            {
                                data.Add(single_loan_detail_data[0]);
                            }
                            else
                            {
                                single_loan_detail_data.Add( System.DateTime.Now.ToString()+ ",,,,,,,,,無資料,,,,,,,,,無資料,,,,,,,,,無資料,,,,,,,,,無資料"); ////////////需要檢查一下
                                data.Add(single_loan_detail_data[0]);
                            }
                        }
                        data.Insert(0, string.Format("成功獲取「債務」資料。\r\n"));
                        function_data_process.Set_Customer_Loan_Detail_Data(ref temp_list_pre_customers, data, j, i);

                    }
                }
            }
            catch (Exception)
            {

                data.Insert(0, string.Format("獲取「債務」失敗，請聯絡【研發中心-郁宸】。\r\n"));
            }

        }

        public void Load_Building_Base_Data()
        {
            try
            {
                for (int i = 0; i <= temp_list_pre_customers.Count - 1; i++)
                {
                    string value = temp_list_pre_customers[i].buildingbaseId;
                    data = controller_sql_categoryt.Controller_SQL_Action_Category(Index_SQL_Action_Category.Get, Index_SQL_Action_Function.Get_Building_Type, conn, value);
                    data.Insert(0, string.Format("成功獲取「建物類型」資料。\r\n"));
                    function_data_process.Set_Building_Base_Data(ref temp_list_pre_customers, data, i);
                }
            }
            catch (Exception)
            {
                data.Insert(0, string.Format("獲取「建物類型」資料失敗，請聯絡【研發中心-郁宸】。\r\n"));
            }
            /// 地區更新

        }

        public void Load_Updating_Time_Data()
        {
            /// 紀錄更新時間
            /// 資料內容有：【更新時間】
            for (int i = 0; i <= temp_list_pre_customers.Count - 1; i++)
            {
                temp_list_pre_customers[i].updating_time = DateTime.Now.ToString("yyyy/MM/dd-hh:mm:ss");
            }
        }





        #region 測試區
        public string Test_CADVVISER()
        {
            set_sql_conn.Load_SQLConn_Str_CADVVISER_DB(); // Load CADVISER_DB Conn Data
            SqlConnection conn = new SqlConnection(set_sql_conn.conn_str);

           return  controller_sql_categoryt.Controller_SQL_Action_Category(Index_SQL_Action_Category.Test, Index_SQL_Action_Function.Try_Connection, conn, set_sql_conn.Data_Base)[0];
        }
        public string Test_Temp_DB()
        {
            set_sql_conn.Load_SQLConn_Str_Temp_DB(); // Load CADVISER_DB Conn Data
            SqlConnection conn = new SqlConnection(set_sql_conn.conn_str);

            return controller_sql_categoryt.Controller_SQL_Action_Category(Index_SQL_Action_Category.Test, Index_SQL_Action_Function.Try_Connection, conn, set_sql_conn.Data_Base)[0];
        }
        #endregion
    }


    class Function_Evaluate_Data
    {
        List<Set_Pre_Customer_Data> temp_list_pre_customers = new List<Set_Pre_Customer_Data>();
        Function_Evaluate function_evaluate = new Function_Evaluate();
        public void Set_pre_customers_data(List<Set_Pre_Customer_Data> list_pre_customers)
        {
            temp_list_pre_customers = list_pre_customers;
        }


        public void Evaluate_Score()
        {
            Evaluate_Rigester_Date(); /// 評量登記時間
            Evaluate_Loan_Order_Gap(); /// 評量他項登記次序的差值
            Evaluate_Range_Part(); /// 評量權力範圍
            Evaluate_Register_Reason(); /// 評量登記原因
            Evaluate_Area(); /// 評量物件面積
            Evaluate_Floor(); /// 評量物件樓數
            Evaluate_First_Name(); /// 評量債權人的姓氏
            Sum_Evaluate_Result(); /// 計算總評量分數
        }
        public void Evaluate_Rigester_Date()
        {
            temp_list_pre_customers = function_evaluate.Register_Date(temp_list_pre_customers);
        }
        public void Evaluate_Loan_Order_Gap()
        {
            temp_list_pre_customers = function_evaluate.Loan_Order_Gap(temp_list_pre_customers);
        }
        public void Evaluate_Range_Part()
        {
            temp_list_pre_customers = function_evaluate.Holder_Area_Part(temp_list_pre_customers);
        }
        public void Evaluate_Register_Reason()
        {
            temp_list_pre_customers = function_evaluate.Register_Reason(temp_list_pre_customers);
        }
        public void Evaluate_Area()
        {
            temp_list_pre_customers = function_evaluate.Total_Area(temp_list_pre_customers);
        }
        public void Evaluate_Floor()
        {
            temp_list_pre_customers = function_evaluate.Total_Floor(temp_list_pre_customers);
        }
        public void Evaluate_First_Name()
        {
            temp_list_pre_customers = function_evaluate.Evaluate_Droit_Rigester_First_Name(temp_list_pre_customers);
        }

        public void Sum_Evaluate_Result()
        {
            temp_list_pre_customers = function_evaluate.Sum_Total_Score(temp_list_pre_customers);
        }

    }

    class Function_Export_Excel
    {
        List<Set_Pre_Customer_Data> temp_list_pre_customers = new List<Set_Pre_Customer_Data>();
        Function_Excel function_excel = new Function_Excel();
        
        public void Export_Excel(List<Set_Pre_Customer_Data> list_pre_customers)
        {
            //var sortedList = list_pre_customers.OrderByDescending(x => x.).ToList();

            function_excel.Export_Excel(list_pre_customers);
        }
    }
}
