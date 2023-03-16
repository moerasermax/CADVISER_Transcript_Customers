using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using CADVISER_Transcript_Customers.Model;
using CADVISER_Transcript_Customers.Model.Data_Set.Index_Set;
using CADVISER_Transcript_Customers.Model.Data_Set;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace CADVISER_Transcript_Customers.Model
{
    public class Function_DataProcess
    {
        #region SQL區
        public List<string> Process_SQL_Result(Index_SQL_Action_Function query_str, DataTable dataTable)
        {
            List<string> result = new List<string>();

            try
            {
                switch (query_str)
                {

                    case Index_SQL_Action_Function.Get_Building:
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            string return_str = dr["BuildingBaseId"].ToString()+",,,,,,,,";
                            return_str += dr["建物門牌"].ToString() + ",,,,,,,,";
                            return_str += dr["總面積"].ToString() + ",,,,,,,,";
                            return_str += dr["層數"].ToString();
                            
                            result.Add(return_str);
                        }
                        break;
                    case Index_SQL_Action_Function.Get_Owner:
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            string return_str = "";
                            return_str += dr["所有權人姓名"].ToString() + ",,,,,,,,,";
                            return_str += dr["統一編號"].ToString() + ",,,,,,,,,";
                            return_str += dr["登記原因"].ToString() + ",,,,,,,,,";
                            return_str += dr["權利範圍"].ToString() + ",,,,,,,,,";
                            return_str += dr["相關他項登記次序"].ToString();

                            result.Add(return_str);

                        }
                        break;
                    case Index_SQL_Action_Function.Get_Loan_Detail:
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            string return_str = "";
                            return_str += dr["登記日期"].ToString() + ",,,,,,,,,";
                            return_str += dr["權利人姓名"].ToString() + ",,,,,,,,,";
                            return_str += dr["擔保債權總金額"].ToString()+ ",,,,,,,,,";
                            return_str += dr["他項登記次序"].ToString();
                            result.Add(return_str);
                        }
                        break;
                    case Index_SQL_Action_Function.Get_Building_Type:
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            string return_str = "";
                            return_str += dr["City"].ToString() + ",,,,,,,,,";
                            return_str += dr["Township"].ToString() + ",,,,,,,,,";
                            return_str += dr["建物類型"].ToString() + ",,,,,,,,,";
                            return_str += dr["Section"].ToString() + ",,,,,,,,,";
                            return_str += dr["BuildingNumber"].ToString() ;
                            result.Add(return_str);
                        }
                        break;
                    default:
                        result.Add("資料無法解析，請聯絡【研發中心-郁宸】。");
                        break;
                }



            }
            catch (Exception)
            {

                result.Add("資料解析發生錯誤，請聯絡【研發中心-郁宸】。");
            }


            return result;
        }

        #endregion

        #region 記憶體區
        public void Set_BuildingBaseID(ref List<Set_Pre_Customer_Data> list_pre_customer_data,List<string> data)
        {
            /// 因為資料庫欄位是中文，不能直接用 class接
            for (int i = 1; i <= data.Count -1; i++)
            {
                if (i == 936) { }
                string[] data_arr = Regex.Split(data[i], ",,,,,,,,");
                Set_Pre_Customer_Data single_pre_customer_data = new Set_Pre_Customer_Data();
                
                single_pre_customer_data.buildingbaseId = data_arr[0];
                single_pre_customer_data.information_address = data_arr[1];
                single_pre_customer_data.information_total_area = data_arr[2].Replace(")","").Replace("(","-").Replace("平方公尺","").Replace("坪",""); /// 資料格式：平方公尺-坪數
                single_pre_customer_data.information_total_floor = data_arr[3].Replace("層","");

                list_pre_customer_data.Add(single_pre_customer_data);
                
            }


                

            
        }

        public void Set_Customer_Data(ref List<Set_Pre_Customer_Data> list_pre_customer_data, List<string> data,int list_pre_customer_index)
        {
            List<Information_Customer> information_data = new List<Information_Customer>();
            int count = 0;
            for (int i = 1; i <= data.Count-1; i++)

            {
                Information_Customer single_information_data = new Information_Customer();
                information_data.Add(single_information_data);
                string[] data_arr = Regex.Split(data[i], ",,,,,,,,,");
                information_data[count].name = data_arr[0];
                information_data[count].idnetity = data_arr[1];
                information_data[count].droit_source = data_arr[2];
                information_data[count].holding_range = data_arr[3].Replace("全部","").Replace(" ","").Replace("分之","/"); //格式：分母/子
                information_data[count].loan_tittle = data_arr[4];
                if (information_data[count].idnetity.Substring(1, 1).Equals("2"))
                {
                    information_data[count].gender = "女";
                }
                else
                {
                    information_data[count].gender = "男";
                }
                count++;
            }

            list_pre_customer_data[list_pre_customer_index].customer_data = information_data;



        }

        public void Set_Customer_Loan_Detail_Data(ref List<Set_Pre_Customer_Data> list_pre_customer_data , List<string> data,int customer_index,int list_pre_customer_data_index)
        {
            List<Information__Loan_detail> list_loan_details = new List<Information__Loan_detail>();
            for (int i = 1; i <= data.Count -1  ; i++)
            {
                string[] data_arr = Regex.Split(data[i], ",,,,,,,,,");


                Information__Loan_detail single_loan_details = new Information__Loan_detail();

                single_loan_details.register_date = data_arr[0];
                single_loan_details.Loan_Authority = data_arr[1];
                single_loan_details.Loan_Amount = data_arr[2].Replace(" ","").Replace("新台幣","").Replace("元正", "");
                single_loan_details.order_no = data_arr[3];

                list_loan_details.Add(single_loan_details);
            }
            list_pre_customer_data[list_pre_customer_data_index].customer_data[customer_index].loan_details = list_loan_details;

        }

        public void Set_Building_Base_Data(ref List<Set_Pre_Customer_Data> list_pre_customer_data, List<string> data,int list_pre_customer_data_index)
        {
            for (int i = 1; i <= data.Count -1 ; i++)
            {
                string[] data_arr = Regex.Split(data[i], ",,,,,,,,,");

                list_pre_customer_data[list_pre_customer_data_index].information_address = data_arr[0] + data_arr[1] + list_pre_customer_data[list_pre_customer_data_index].information_address;
                list_pre_customer_data[list_pre_customer_data_index].information_building_type = data_arr[2];
                list_pre_customer_data[list_pre_customer_data_index].information_building_section = data_arr[3];
                list_pre_customer_data[list_pre_customer_data_index].information_building_buildingnumber = data_arr[4];
                list_pre_customer_data[list_pre_customer_data_index].information_district = data_arr[0] + data_arr[1];
            }
        }

        #endregion

    }
}
