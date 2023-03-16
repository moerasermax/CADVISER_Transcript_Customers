using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using CADVISER_Transcript_Customers.Model.Data_Set;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;
using Excel = Microsoft.Office.Interop.Excel;
using Loading_Bar = Loadding_Bar;

namespace CADVISER_Transcript_Customers.Model
{
    public class Function_Excel
    {

        public void Export_Excel(List<Set_Pre_Customer_Data> list_pre_customers)
        {
            string file_name = DateTime.Now.Hour + "_" + DateTime.Now.Minute + "_" + DateTime.Now.Second;
            string Desktop_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\客戶謄本分析";
            string Folder_path = Desktop_path + "\\" + DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString() + DateTime.Today.Day.ToString();
            string File_name = DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString();

            if (!Directory.Exists(Folder_path))
            {
                Directory.CreateDirectory(Folder_path);
            }




            string File_str = Folder_path + "\\" + File_name + "_" + file_name;
            Excel.Application Excel_App = new Excel.Application();  // Excel 應用程式啟動
            Excel.Workbook Excel_WB = Excel_App.Workbooks.Add(true); // 工作區建立
            Excel.Worksheet Excel_Error_Sheet = Excel_WB.ActiveSheet as Excel.Worksheet; // 活頁簿建立 // 記錄錯誤用
            Excel.Worksheet Excel_Reference_WS = new Excel.Worksheet();
            Excel_Error_Sheet.Name = "資料無法解析";

            int Error_index = 1; // 記錄錯誤用
            Excel_Error_Sheet.Cells[Error_index] = "建物ID";
            Excel_Error_Sheet.Cells[Error_index] = "客戶唯一編號";

            Sheet_1(Excel_App, Excel_WB ,list_pre_customers, Excel_Error_Sheet, ref Error_index);
            Sheet_2(Excel_App, Excel_WB, list_pre_customers, Excel_Error_Sheet, ref Error_index);
            Sheet_3(Excel_App, Excel_WB, list_pre_customers, Excel_Error_Sheet, ref Error_index, ref Excel_Reference_WS);


            /// Level_A 根據分數做排行分群
            Sheet_Level_A_New(Excel_WB, Excel_Reference_WS);
            Sheet_Level_B_New(Excel_WB, Excel_Reference_WS);
            Sheet_Level_C_New(Excel_WB, Excel_Reference_WS);

            //Sheet_Level_A(Excel_App ,Excel_WB, list_pre_customers);
            //Sheet_Level_B(Excel_App, Excel_WB, list_pre_customers);
            //Sheet_Level_C(Excel_App, Excel_WB, list_pre_customers);
            //Sheet_Level_D(Excel_App, Excel_WB, list_pre_customers);
            Excel_WB.SaveAs(File_str);
            Excel_WB.Close(false, Missing.Value, Missing.Value);
            Excel_App.Quit();

            Excel_WB = null;
            Excel_App = null;
            GC.Collect();

            MessageBox.Show("Excel 輸出成功");

        }


        public void Sheet_1(Excel.Application Excel_App, Excel.Workbook Excel_WB, List<Set_Pre_Customer_Data> list_pre_customers,Excel.Worksheet Excel_Error_Sheet,ref int Error_index)
        {
            Excel.Worksheet Excel_WS = Excel_WB.ActiveSheet as Excel.Worksheet;
            Excel_WS = Excel_WB.Worksheets[1];
            Excel_WS.Name = "分析結果";
            Set_Font_Style(ref Excel_WS);
            Set_Tittle_Freeze(ref Excel_WS);

            Excel_WS.Cells[1, 1] = "項次";
            Excel_WS.Cells[1, 2] = "評分";
            Excel_WS.Cells[1, 3] = "客戶姓名";
            Excel_WS.Cells[1, 4] = "性別";
            Excel_WS.Cells[1, 5] = "身份證號";
            Excel_WS.Cells[1, 6] = "登記原因";
            Excel_WS.Cells[1, 7] = "標的物地址 / 所有權地址";
            Excel_WS.Cells[1, 8] = "建物類型";
            Excel_WS.Cells[1, 9] = "建物坪㎡";
            Excel_WS.Cells[1, 10] = "建物坪";
            Excel_WS.Cells[1, 11] = "一胎時間";
            Excel_WS.Cells[1, 12] = "一胎債權";
            Excel_WS.Cells[1, 13] = "一胎設定債權(萬)";
            Excel_WS.Cells[1, 14] = "二胎時間";
            Excel_WS.Cells[1, 15] = "二胎債權";
            Excel_WS.Cells[1, 16] = "二胎設定債權(萬)";
            Excel_WS.Cells[1, 17] = "備註";
            Excel_WS.Cells[1, 18] = "多胎以上資料（時間、債權、設定債權金額";
            Excel_WS.Cells[1, 19] = "跑謄狀態";
            Excel_WS.Cells[1, 20] = "電話";
            Excel_WS.Cells[1, 21] = "聯徵";
            Excel_WS.Cells[1, 22] = "開發日期";
            Excel_WS.Cells[1, 23] = "段號_建號";
            Excel_WS.Cells[1, 24] = "區域";

            string range_str = string.Format("A{0}:X{1}", 1.ToString(), 1.ToString());
            Set_Tittle_Column_Style(ref Excel_WS, range_str);
            try
            {
                string Current_BuildBaseID_CustomerIdentity = ""; // 記錄錯誤用
                int index = 1;
                Loadding_Bar.Create_Loadding_UI loading_bar = new Loading_Bar.Create_Loadding_UI();////////////
                loading_bar.Set_Maximum(list_pre_customers.Count - 1);//////////////////

                for (int i = 0; i <= list_pre_customers.Count - 1; i++)
                {

                    loading_bar.Update_Loadding_UI(i+5);/////////////

                    for (int j = 0; j <= list_pre_customers[i].customer_data.Count - 1; j++)
                    {
                        Current_BuildBaseID_CustomerIdentity = list_pre_customers[i].buildingbaseId + "_" + list_pre_customers[i].customer_data[j].idnetity;// 記錄錯誤用
                        index += 1;
                        string loan_more_data = "";
                        
                        for (int x = 0; x <= list_pre_customers[i].customer_data[j].loan_details.Count - 1; x++)
                        {
                            try
                            {


                                string[] area_arr = list_pre_customers[i].information_total_area.Split('-');
                                Excel_WS.Cells[index, 1] = (index-1).ToString();
                                Excel_WS.Cells[index, 2] = list_pre_customers[i].customer_data[j].evaluate_Score;
                                Excel_WS.Cells[index, 3] = list_pre_customers[i].customer_data[j].name;
                                Excel_WS.Cells[index, 4] = list_pre_customers[i].customer_data[j].gender;
                                Excel_WS.Cells[index, 5] = list_pre_customers[i].customer_data[j].idnetity;
                                Excel_WS.Cells[index, 6] = list_pre_customers[i].customer_data[j].droit_source;


                                Excel_WS.Cells[index, 7] = list_pre_customers[i].information_address;
                                Excel_WS.Cells[index, 8] = list_pre_customers[i].information_building_type;
                                Excel_WS.Cells[index, 9] = area_arr[0];
                                Excel_WS.Cells[index, 10] = area_arr[1];

                                if (x == 0)
                                {
                                    Excel_WS.Cells[index, 11] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                    if(list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                    {
                                        Excel_WS.Cells[index, 12] = "無法解析此債權人姓名";
                                    }
                                    else
                                    {
                                        Excel_WS.Cells[index, 12] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                    }
                                    Excel_WS.Cells[index, 13] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                }
                                if (x == 1)
                                {
                                    Excel_WS.Cells[index, 14] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                    if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                    {
                                        Excel_WS.Cells[index, 15] = "無法解析此債權人姓名";
                                    }
                                    else
                                    {
                                        Excel_WS.Cells[index, 15] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                    }
                                    Excel_WS.Cells[index, 16] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                }
                                Excel_WS.Cells[index, 17] = "";
                                if (x >= 2)
                                {
                                    loan_more_data += "-----第" + (x + 1).ToString() + "胎----\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].register_date + "\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority + "\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount + "\r\n";
                                    Excel_WS.Cells[index, 18] = loan_more_data;
                                }
                                Excel_WS.Cells[index, 19] = "";
                                Excel_WS.Cells[index, 20] = "";
                                Excel_WS.Cells[index, 21] = "";
                                Excel_WS.Cells[index, 22] = "";
                                Excel_WS.Cells[index, 23] = list_pre_customers[i].information_building_section + "_" + list_pre_customers[i].information_building_buildingnumber;
                                Excel_WS.Cells[index, 24] = list_pre_customers[i].information_district;

                                range_str = string.Format("A{0}:X{1}", index, index);
                                Set_Column_Style(ref Excel_WS, range_str);
                            }
                            catch (Exception ex)
                            {
                                Recording_Error_Data(Excel_Error_Sheet, ref Error_index, Current_BuildBaseID_CustomerIdentity);
                            }
                        }

                    }
                }
                Excel_WS.Range["A1", "Z" + index].RowHeight = 35;
                Sort(ref Excel_WS, index.ToString(), "B," + "X," + "District");
                Excel_WS = null;
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message.ToString());
            }
        }
        public void Sheet_2(Excel.Application Excel_App, Excel.Workbook Excel_WB, List<Set_Pre_Customer_Data> list_pre_customers, Excel.Worksheet Excel_Error_Sheet, ref int Error_index)
        {
            object missing = Type.Missing;
            Excel.Worksheet Excel_WS = Excel_WB.Sheets.Add(missing, missing, 1, missing) as Excel.Worksheet;
            Excel_WS.Name = "評量細項_區域排行";
            Set_Font_Style(ref Excel_WS);
            Set_Tittle_Freeze(ref Excel_WS);

            Excel_WS.Cells[1, 1] = "項次";
            Excel_WS.Cells[1, 2] = "客戶姓名";
            Excel_WS.Cells[1, 3] = "性別";
            Excel_WS.Cells[1, 4] = "身份證號";
            Excel_WS.Cells[1, 5] = "標的物地址 / 所有權地址標的物地址 / 所有權地址";
            Excel_WS.Cells[1, 6] = "建物類型";
            Excel_WS.Cells[1, 7] = "建物坪";
            Excel_WS.Cells[1, 8] = "一胎時間";
            Excel_WS.Cells[1, 9] = "一胎債權";
            Excel_WS.Cells[1, 10] = "一胎設定債權(萬)";
            Excel_WS.Cells[1, 11] = "二胎時間";
            Excel_WS.Cells[1, 12] = "二胎債權";
            Excel_WS.Cells[1, 13] = "二胎設定債權(萬)";
            Excel_WS.Cells[1, 14] = "備註";
            Excel_WS.Cells[1, 15] = "多胎以上資料（時間、債權、設定債權金額)";
            Excel_WS.Cells[1, 16] = "評分";
            Excel_WS.Cells[1, 17] = "登記原因";
            Excel_WS.Cells[1, 18] = "評分項目-登記原因";
            Excel_WS.Cells[1, 19] = "建物㎡";
            Excel_WS.Cells[1, 20] = "評分項目-建物㎡";
            Excel_WS.Cells[1, 21] = "他項登記次序";
            Excel_WS.Cells[1, 22] = "評分項目-他項登記次序";
            Excel_WS.Cells[1, 23] = "登記日期";
            Excel_WS.Cells[1, 24] = "評分項目-登記日期";
            Excel_WS.Cells[1, 25] = "權力範圍";
            Excel_WS.Cells[1, 26] = "評分項目-權力範圍";
            Excel_WS.Cells[1, 27] = "建物總樓層";
            Excel_WS.Cells[1, 28] = "評分項目-建物總樓層";
            Excel_WS.Cells[1, 29] = "權力人姓名";
            Excel_WS.Cells[1, 30] = "評分項目-權力人姓名";
            Excel_WS.Cells[1, 31] = "跑謄狀態";
            Excel_WS.Cells[1, 32] = "電話";
            Excel_WS.Cells[1, 33] = "聯徵";
            Excel_WS.Cells[1, 34] = "開發日期";
            Excel_WS.Cells[1, 35] = "段號_建號";
            Excel_WS.Cells[1, 36] = "區域";

            string range_str = string.Format("A{0}:AJ{1}", 1.ToString(), 1.ToString());
            Set_Tittle_Column_Style(ref Excel_WS, range_str);
            try
            {
                string Current_BuildBaseID_CustomerIdentity = ""; // 記錄錯誤用
                int index = 1;
                Loadding_Bar.Create_Loadding_UI loading_bar = new Loading_Bar.Create_Loadding_UI();////////////
                loading_bar.Set_Maximum(list_pre_customers.Count - 1);//////////////////

                for (int i = 0; i <= list_pre_customers.Count - 1; i++)
                {
                    loading_bar.Update_Loadding_UI(i + 5);/////////////


                    for (int j = 0; j <= list_pre_customers[i].customer_data.Count - 1; j++)
                    {
                        Current_BuildBaseID_CustomerIdentity = list_pre_customers[i].buildingbaseId + "_" + list_pre_customers[i].customer_data[j].idnetity;// 記錄錯誤用

                        if (list_pre_customers[i].customer_data[j].loan_details[0].Loan_Authority.Length <= 30)
                        {

                            index += 1;
                            string loan_more_data = "";
                            for (int x = 0; x <= list_pre_customers[i].customer_data[j].loan_details.Count - 1; x++)
                            {
                                try
                                {


                                    string[] area_arr = list_pre_customers[i].information_total_area.Split('-');
                                    Excel_WS.Cells[index, 1] = index.ToString();
                                    Excel_WS.Cells[index, 2] = list_pre_customers[i].customer_data[j].name;
                                    Excel_WS.Cells[index, 3] = list_pre_customers[i].customer_data[j].gender;
                                    Excel_WS.Cells[index, 4] = list_pre_customers[i].customer_data[j].idnetity;
                                    Excel_WS.Cells[index, 5] = list_pre_customers[i].information_address;
                                    Excel_WS.Cells[index, 6] = list_pre_customers[i].information_building_type;
                                    Excel_WS.Cells[index, 7] = area_arr[0];
                                    if (x == 0)
                                    {
                                        Excel_WS.Cells[index, 8] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                        if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                        {
                                            Excel_WS.Cells[index, 9] = "無法解析此債權人姓名";
                                        }
                                        else
                                        {
                                            Excel_WS.Cells[index, 9] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                        }
                                        Excel_WS.Cells[index, 10] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                    }
                                    if (x == 1)
                                    {
                                        Excel_WS.Cells[index, 11] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                        if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                        {
                                            Excel_WS.Cells[index, 12] = "無法解析此債權人姓名";
                                        }
                                        else
                                        {
                                            Excel_WS.Cells[index, 12] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                        }
                                        Excel_WS.Cells[index, 13] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                    }
                                    Excel_WS.Cells[index, 14] = "";
                                    if (x >= 2)
                                    {
                                        loan_more_data += "-----第" + (x + 1).ToString() + "胎----\r\n";
                                        loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].register_date + "\r\n";
                                        loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority + "\r\n";
                                        loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount + "\r\n";
                                        Excel_WS.Cells[index, 15] = loan_more_data;
                                    }

                                    Excel_WS.Cells[index, 16] = list_pre_customers[i].customer_data[j].evaluate_Score.ToString();
                                    Excel_WS.Cells[index, 17] = list_pre_customers[i].customer_data[j].droit_source;
                                    Excel_WS.Cells[index, 18] = list_pre_customers[i].customer_data[j].evaluate_droit_source;
                                    Excel_WS.Cells[index, 19] = area_arr[0];
                                    Excel_WS.Cells[index, 20] = list_pre_customers[i].evaluate_total_area;
                                    Excel_WS.Cells[index, 21] = list_pre_customers[i].customer_data[j].loan_tittle;
                                    Excel_WS.Cells[index, 22] = list_pre_customers[i].customer_data[j].evaluate_loan_order_gap;
                                    Excel_WS.Cells[index, 23] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                    Excel_WS.Cells[index, 24] = list_pre_customers[i].customer_data[j].evaluate_rigester_Date;
                                    Excel_WS.Cells[index, 25].NumberFormat = "@";
                                    Excel_WS.Cells[index, 25] = list_pre_customers[i].customer_data[j].holding_range;
                                    Excel_WS.Cells[index, 26] = list_pre_customers[i].customer_data[j].evaluate_holder_area_part;
                                    Excel_WS.Cells[index, 27] = list_pre_customers[i].information_total_floor;
                                    Excel_WS.Cells[index, 28] = list_pre_customers[i].evaluate_total_floor;
                                    Excel_WS.Cells[index, 29] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                    Excel_WS.Cells[index, 30] = list_pre_customers[i].customer_data[j].evaluate_droit_rigester_first_name;
                                    Excel_WS.Cells[index, 31] = "";
                                    Excel_WS.Cells[index, 32] = "";
                                    Excel_WS.Cells[index, 33] = "";
                                    Excel_WS.Cells[index, 34] = "";
                                    Excel_WS.Cells[index, 35] = list_pre_customers[i].information_building_section + "_" + list_pre_customers[i].information_building_buildingnumber;
                                    Excel_WS.Cells[index, 36] = list_pre_customers[i].information_district;

                                    range_str = string.Format("A{0}:AJ{1}", index, index);
                                    Set_Column_Style(ref Excel_WS, range_str);


                                }
                                catch (Exception)
                                {
                                    Recording_Error_Data(Excel_Error_Sheet, ref Error_index, Current_BuildBaseID_CustomerIdentity);
                                }
                            }
                        }
                    }
                }

                Excel_WS.Range["A1", "Z" + index].RowHeight = 35;
                Sort(ref Excel_WS, index.ToString(), "P," + "AJ," + "District");
                Excel_WS = null;

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message.ToString());
            }
        }
        public void Sheet_3(Excel.Application Excel_App, Excel.Workbook Excel_WB, List<Set_Pre_Customer_Data> list_pre_customers, Excel.Worksheet Excel_Error_Sheet, ref int Error_index, ref Excel.Worksheet Excel_Reference_WS)
        {
            object missing = Type.Missing;
            Excel.Worksheet Excel_WS = Excel_WB.Sheets.Add(missing, missing, 1, missing) as Excel.Worksheet;
            Excel_WS.Name = "評量細項_評分排行";


            Excel_WS.Cells[1, 1] = "項次";
            Excel_WS.Cells[1, 2] = "客戶姓名";
            Excel_WS.Cells[1, 3] = "性別";
            Excel_WS.Cells[1, 4] = "身份證號";
            Excel_WS.Cells[1, 5] = "標的物地址 / 所有權地址標的物地址 / 所有權地址";
            Excel_WS.Cells[1, 6] = "建物類型";
            Excel_WS.Cells[1, 7] = "建物坪";
            Excel_WS.Cells[1, 8] = "一胎時間";
            Excel_WS.Cells[1, 9] = "一胎債權";
            Excel_WS.Cells[1, 10] = "一胎設定債權(萬)";
            Excel_WS.Cells[1, 11] = "二胎時間";
            Excel_WS.Cells[1, 12] = "二胎債權";
            Excel_WS.Cells[1, 13] = "二胎設定債權(萬)";
            Excel_WS.Cells[1, 14] = "備註";
            Excel_WS.Cells[1, 15] = "多胎以上資料（時間、債權、設定債權金額)";
            Excel_WS.Cells[1, 16] = "評分";
            Excel_WS.Cells[1, 17] = "登記原因";
            Excel_WS.Cells[1, 18] = "評分項目-登記原因";
            Excel_WS.Cells[1, 19] = "建物㎡";
            Excel_WS.Cells[1, 20] = "評分項目-建物㎡";
            Excel_WS.Cells[1, 21] = "他項登記次序";
            Excel_WS.Cells[1, 22] = "評分項目-他項登記次序";
            Excel_WS.Cells[1, 23] = "登記日期";
            Excel_WS.Cells[1, 24] = "評分項目-登記日期";
            Excel_WS.Cells[1, 25] = "權力範圍";
            Excel_WS.Cells[1, 26] = "評分項目-權力範圍";
            Excel_WS.Cells[1, 27] = "建物總樓層";
            Excel_WS.Cells[1, 28] = "評分項目-建物總樓層";
            Excel_WS.Cells[1, 29] = "權力人姓名";
            Excel_WS.Cells[1, 30] = "評分項目-權力人姓名";
            Excel_WS.Cells[1, 31] = "跑謄狀態";
            Excel_WS.Cells[1, 32] = "電話";
            Excel_WS.Cells[1, 33] = "聯徵";
            Excel_WS.Cells[1, 34] = "開發日期";
            Excel_WS.Cells[1, 35] = "段號_建號";
            Excel_WS.Cells[1, 36] = "區域";

            string range_str = string.Format("A{0}:AJ{1}", 1.ToString(), 1.ToString());
            Set_Tittle_Column_Style(ref Excel_WS, range_str);
            Set_Font_Style(ref Excel_WS);
            Set_Tittle_Freeze(ref Excel_WS);

            try
            {
                string Current_BuildBaseID_CustomerIdentity = ""; // 記錄錯誤用
                int index = 1;
                Loadding_Bar.Create_Loadding_UI loading_bar = new Loading_Bar.Create_Loadding_UI();////////////
                loading_bar.Set_Maximum(list_pre_customers.Count - 1);//////////////////

                for (int i = 0; i <= list_pre_customers.Count - 1; i++)
                {
                    loading_bar.Update_Loadding_UI(i + 5);/////////////


                    for (int j = 0; j <= list_pre_customers[i].customer_data.Count - 1; j++)
                    {
                        Current_BuildBaseID_CustomerIdentity = list_pre_customers[i].buildingbaseId + "_" + list_pre_customers[i].customer_data[j].idnetity;// 記錄錯誤用

                        if (list_pre_customers[i].customer_data[j].loan_details[0].Loan_Authority.Length <= 30)
                        {

                            index += 1;
                            string loan_more_data = "";
                            for (int x = 0; x <= list_pre_customers[i].customer_data[j].loan_details.Count - 1; x++)
                            {
                                try
                                {


                                    string[] area_arr = list_pre_customers[i].information_total_area.Split('-');
                                    Excel_WS.Cells[index, 1] = index.ToString();
                                    Excel_WS.Cells[index, 2] = list_pre_customers[i].customer_data[j].name;
                                    Excel_WS.Cells[index, 3] = list_pre_customers[i].customer_data[j].gender;
                                    Excel_WS.Cells[index, 4] = list_pre_customers[i].customer_data[j].idnetity;
                                    Excel_WS.Cells[index, 5] = list_pre_customers[i].information_address;
                                    Excel_WS.Cells[index, 6] = list_pre_customers[i].information_building_type;
                                    Excel_WS.Cells[index, 7] = area_arr[0];
                                    if (x == 0)
                                    {
                                        Excel_WS.Cells[index, 8] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                        if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                        {
                                            Excel_WS.Cells[index, 9] = "無法解析此債權人姓名";
                                        }
                                        else
                                        {
                                            Excel_WS.Cells[index, 9] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                        }
                                        Excel_WS.Cells[index, 10] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                    }
                                    if (x == 1)
                                    {
                                        Excel_WS.Cells[index, 11] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                        if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                        {
                                            Excel_WS.Cells[index, 12] = "無法解析此債權人姓名";
                                        }
                                        else
                                        {
                                            Excel_WS.Cells[index, 12] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                        }
                                        Excel_WS.Cells[index, 13] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                    }
                                    Excel_WS.Cells[index, 14] = "";
                                    if (x >= 2)
                                    {
                                        loan_more_data += "-----第" + (x + 1).ToString() + "胎----\r\n";
                                        loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].register_date + "\r\n";
                                        loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority + "\r\n";
                                        loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount + "\r\n";
                                        Excel_WS.Cells[index, 15] = loan_more_data;
                                    }

                                    Excel_WS.Cells[index, 16] = list_pre_customers[i].customer_data[j].evaluate_Score.ToString();
                                    Excel_WS.Cells[index, 17] = list_pre_customers[i].customer_data[j].droit_source;
                                    Excel_WS.Cells[index, 18] = list_pre_customers[i].customer_data[j].evaluate_droit_source;
                                    Excel_WS.Cells[index, 19] = area_arr[0];
                                    Excel_WS.Cells[index, 20] = list_pre_customers[i].evaluate_total_area;
                                    Excel_WS.Cells[index, 21] = list_pre_customers[i].customer_data[j].loan_tittle;
                                    Excel_WS.Cells[index, 22] = list_pre_customers[i].customer_data[j].evaluate_loan_order_gap;
                                    Excel_WS.Cells[index, 23] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                    Excel_WS.Cells[index, 24] = list_pre_customers[i].customer_data[j].evaluate_rigester_Date;
                                    Excel_WS.Cells[index, 25].NumberFormat = "@";
                                    Excel_WS.Cells[index, 25] = list_pre_customers[i].customer_data[j].holding_range;
                                    Excel_WS.Cells[index, 26] = list_pre_customers[i].customer_data[j].evaluate_holder_area_part;
                                    Excel_WS.Cells[index, 27] = list_pre_customers[i].information_total_floor;
                                    Excel_WS.Cells[index, 28] = list_pre_customers[i].evaluate_total_floor;
                                    Excel_WS.Cells[index, 29] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                    Excel_WS.Cells[index, 30] = list_pre_customers[i].customer_data[j].evaluate_droit_rigester_first_name;
                                    Excel_WS.Cells[index, 31] = "";
                                    Excel_WS.Cells[index, 32] = "";
                                    Excel_WS.Cells[index, 33] = "";
                                    Excel_WS.Cells[index, 34] = "";
                                    Excel_WS.Cells[index, 35] = list_pre_customers[i].information_building_section + "_" + list_pre_customers[i].information_building_buildingnumber;
                                    Excel_WS.Cells[index, 36] = list_pre_customers[i].information_district;

                                    range_str = string.Format("A{0}:AJ{1}", index, index);
                                    Set_Column_Style(ref Excel_WS, range_str);


                                }
                                catch (Exception)
                                {
                                    Recording_Error_Data(Excel_Error_Sheet, ref Error_index, Current_BuildBaseID_CustomerIdentity);
                                }
                            }
                        }
                    }
                    Excel_Reference_WS = Excel_WS;
                }

                Excel_WS.Range["A1", "Z" + index].RowHeight = 35;
                Sort(ref Excel_WS, index.ToString(), "P," + "AJ," + "Score");
                Excel_WS = null;

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message.ToString());
            }
        }

        public void Sheet_Level_A(Excel.Application Excel_App, Excel.Workbook Excel_WB, List<Set_Pre_Customer_Data> list_pre_customers)
        {
            object missing = Type.Missing;
            Excel.Worksheet Excel_WS = Excel_WB.Sheets.Add(missing, missing, 1, missing) as Excel.Worksheet;
            Excel_WS.Name = "A級名單";
            Excel_WS.Cells.NumberFormat = "@";

            Excel_WS.Cells[1, 1] = "項次";
            Excel_WS.Cells[1, 2] = "客戶姓名";
            Excel_WS.Cells[1, 3] = "性別";
            Excel_WS.Cells[1, 4] = "身份證號";
            Excel_WS.Cells[1, 5] = "標的物地址 / 所有權地址標的物地址 / 所有權地址";
            Excel_WS.Cells[1, 6] = "建物類型";
            Excel_WS.Cells[1, 7] = "建物坪";
            Excel_WS.Cells[1, 8] = "一胎時間";
            Excel_WS.Cells[1, 9] = "一胎債權";
            Excel_WS.Cells[1, 10] = "一胎設定債權(萬)";
            Excel_WS.Cells[1, 11] = "二胎時間";
            Excel_WS.Cells[1, 12] = "二胎債權";
            Excel_WS.Cells[1, 13] = "二胎設定債權(萬)";
            Excel_WS.Cells[1, 14] = "備註";
            Excel_WS.Cells[1, 15] = "多胎以上資料（時間、債權、設定債權金額)";
            Excel_WS.Cells[1, 16] = "評分";
            Excel_WS.Cells[1, 17] = "登記原因";
            Excel_WS.Cells[1, 18] = "評分項目-登記原因";
            Excel_WS.Cells[1, 19] = "建物㎡";
            Excel_WS.Cells[1, 20] = "評分項目-建物㎡";
            Excel_WS.Cells[1, 21] = "他項登記次序";
            Excel_WS.Cells[1, 22] = "評分項目-他項登記次序";
            Excel_WS.Cells[1, 23] = "登記日期";
            Excel_WS.Cells[1, 24] = "評分項目-登記日期";
            Excel_WS.Cells[1, 25] = "權力範圍";
            Excel_WS.Cells[1, 26] = "評分項目-權力範圍";
            Excel_WS.Cells[1, 27] = "建物總樓層";
            Excel_WS.Cells[1, 28] = "評分項目-建物總樓層";
            Excel_WS.Cells[1, 29] = "權力人姓名";
            Excel_WS.Cells[1, 30] = "評分項目-權力人姓名";
            Excel_WS.Cells[1, 31] = "跑謄狀態";
            Excel_WS.Cells[1, 32] = "電話";
            Excel_WS.Cells[1, 33] = "聯徵";
            Excel_WS.Cells[1, 34] = "開發日期";
            Excel_WS.Cells[1, 35] = "段號_建號";

            try
            {
                Loadding_Bar.Create_Loadding_UI loading_bar = new Loading_Bar.Create_Loadding_UI();////////////
                int index = 1;
                loading_bar.Set_Maximum(list_pre_customers.Count - 1);//////////////////

                for (int i = 0; i <= list_pre_customers.Count - 1; i++)
                {
                    
                    loading_bar.Update_Loadding_UI(i + 5);/////////////


                    for (int j = 0; j <= list_pre_customers[i].customer_data.Count - 1; j++)
                    {

                        if (int.Parse(list_pre_customers[i].customer_data[j].evaluate_Score) > 22)
                        {

                        
                        if (list_pre_customers[i].customer_data[j].loan_details[0].Loan_Authority.Length >= 30)
                        {
                            list_pre_customers[i].customer_data[j].loan_details[0].Loan_Authority = "無法解析債權者姓名";
                        }
                            index += 1;
                            string loan_more_data = "";
                            for (int x = 0; x <= list_pre_customers[i].customer_data[j].loan_details.Count - 1; x++)
                            {
                                try
                                {


                                string[] area_arr = list_pre_customers[i].information_total_area.Split('-');
                                Excel_WS.Cells[index, 1] = index.ToString();
                                Excel_WS.Cells[index, 2] = list_pre_customers[i].customer_data[j].name;
                                Excel_WS.Cells[index, 3] = list_pre_customers[i].customer_data[j].gender;
                                Excel_WS.Cells[index, 4] = list_pre_customers[i].customer_data[j].idnetity;
                                Excel_WS.Cells[index, 5] = list_pre_customers[i].information_address;
                                Excel_WS.Cells[index, 6] = list_pre_customers[i].information_building_type;
                                Excel_WS.Cells[index, 7] = area_arr[0];
                                if (x == 0)
                                {
                                    Excel_WS.Cells[index, 8] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                    if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                    {
                                        Excel_WS.Cells[index, 9] = "無法解析此債權人姓名";
                                    }
                                    else
                                    {
                                        Excel_WS.Cells[index, 9] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                    }
                                    Excel_WS.Cells[index, 10] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                }
                                if (x == 1)
                                {
                                    Excel_WS.Cells[index, 11] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                    if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                    {
                                        Excel_WS.Cells[index, 12] = "無法解析此債權人姓名";
                                    }
                                    else
                                    {
                                        Excel_WS.Cells[index, 12] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                    }
                                    Excel_WS.Cells[index, 13] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                }
                                Excel_WS.Cells[index, 14] = "";
                                if (x >= 2)
                                {
                                    loan_more_data += "-----第" + (x + 1).ToString() + "胎----\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].register_date + "\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority + "\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount + "\r\n";
                                    Excel_WS.Cells[index, 15] = loan_more_data;
                                }
                                Excel_WS.Cells[index, 16] = list_pre_customers[i].customer_data[j].evaluate_Score;
                                Excel_WS.Cells[index, 17] = list_pre_customers[i].customer_data[j].droit_source;
                                Excel_WS.Cells[index, 18] = list_pre_customers[i].customer_data[j].evaluate_droit_source;
                                Excel_WS.Cells[index, 19] = area_arr[0];
                                Excel_WS.Cells[index, 20] = list_pre_customers[i].evaluate_total_area;
                                Excel_WS.Cells[index, 21] = list_pre_customers[i].customer_data[j].loan_tittle;
                                Excel_WS.Cells[index, 22] = list_pre_customers[i].customer_data[j].evaluate_loan_order_gap;
                                Excel_WS.Cells[index, 23] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                Excel_WS.Cells[index, 24] = list_pre_customers[i].customer_data[j].evaluate_rigester_Date;
                                Excel_WS.Cells[index, 25] = list_pre_customers[i].customer_data[j].holding_range;
                                Excel_WS.Cells[index, 26] = list_pre_customers[i].customer_data[j].evaluate_holder_area_part;
                                Excel_WS.Cells[index, 27] = list_pre_customers[i].information_total_floor;
                                Excel_WS.Cells[index, 28] = list_pre_customers[i].evaluate_total_floor;
                                Excel_WS.Cells[index, 29] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                Excel_WS.Cells[index, 30] = list_pre_customers[i].customer_data[j].evaluate_droit_rigester_first_name;
                                Excel_WS.Cells[index, 31] = "";
                                Excel_WS.Cells[index, 32] = "";
                                Excel_WS.Cells[index, 33] = "";
                                Excel_WS.Cells[index, 34] = "";
                                Excel_WS.Cells[index, 35] = list_pre_customers[i].information_building_section + "_" + list_pre_customers[i].information_building_buildingnumber;
                                Excel_WS.Cells[index, 36] = list_pre_customers[i].information_district;
                                }
                                catch (Exception)
                                {

                                    throw;
                                }
                            }
                        }
                    }
                }

                Excel_WS.Range["A1", "Z" + index].RowHeight = 35;
                Excel_WS = null;

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message.ToString());
            }
        }
        public void Sheet_Level_B(Excel.Application Excel_App, Excel.Workbook Excel_WB, List<Set_Pre_Customer_Data> list_pre_customers)
        {
            object missing = Type.Missing;
            Excel.Worksheet Excel_WS = Excel_WB.Sheets.Add(missing, missing, 1, missing) as Excel.Worksheet;
            Excel_WS.Name = "B級名單";
            Excel_WS.Cells.NumberFormat = "@";

            Excel_WS.Cells[1, 1] = "項次";
            Excel_WS.Cells[1, 2] = "客戶姓名";
            Excel_WS.Cells[1, 3] = "性別";
            Excel_WS.Cells[1, 4] = "身份證號";
            Excel_WS.Cells[1, 5] = "標的物地址 / 所有權地址標的物地址 / 所有權地址";
            Excel_WS.Cells[1, 6] = "建物類型";
            Excel_WS.Cells[1, 7] = "建物坪";
            Excel_WS.Cells[1, 8] = "一胎時間";
            Excel_WS.Cells[1, 9] = "一胎債權";
            Excel_WS.Cells[1, 10] = "一胎設定債權(萬)";
            Excel_WS.Cells[1, 11] = "二胎時間";
            Excel_WS.Cells[1, 12] = "二胎債權";
            Excel_WS.Cells[1, 13] = "二胎設定債權(萬)";
            Excel_WS.Cells[1, 14] = "備註";
            Excel_WS.Cells[1, 15] = "多胎以上資料（時間、債權、設定債權金額)";
            Excel_WS.Cells[1, 16] = "評分";
            Excel_WS.Cells[1, 17] = "登記原因";
            Excel_WS.Cells[1, 18] = "評分項目-登記原因";
            Excel_WS.Cells[1, 19] = "建物㎡";
            Excel_WS.Cells[1, 20] = "評分項目-建物㎡";
            Excel_WS.Cells[1, 21] = "他項登記次序";
            Excel_WS.Cells[1, 22] = "評分項目-他項登記次序";
            Excel_WS.Cells[1, 23] = "登記日期";
            Excel_WS.Cells[1, 24] = "評分項目-登記日期";
            Excel_WS.Cells[1, 25] = "權力範圍";
            Excel_WS.Cells[1, 26] = "評分項目-權力範圍";
            Excel_WS.Cells[1, 27] = "建物總樓層";
            Excel_WS.Cells[1, 28] = "評分項目-建物總樓層";
            Excel_WS.Cells[1, 29] = "權力人姓名";
            Excel_WS.Cells[1, 30] = "評分項目-權力人姓名";
            Excel_WS.Cells[1, 31] = "跑謄狀態";
            Excel_WS.Cells[1, 32] = "電話";
            Excel_WS.Cells[1, 33] = "聯徵";
            Excel_WS.Cells[1, 34] = "開發日期";
            Excel_WS.Cells[1, 35] = "段號_建號";

            try
            {

                int index = 1;
                Loadding_Bar.Create_Loadding_UI loading_bar = new Loading_Bar.Create_Loadding_UI();////////////
                loading_bar.Set_Maximum(list_pre_customers.Count - 1);//////////////////

                for (int i = 0; i <= list_pre_customers.Count - 1; i++)
                {
                    loading_bar.Update_Loadding_UI(i + 5);/////////////


                    for (int j = 0; j <= list_pre_customers[i].customer_data.Count - 1; j++)
                    {
                        if (22 >= int.Parse(list_pre_customers[i].customer_data[j].evaluate_Score) && int.Parse(list_pre_customers[i].customer_data[j].evaluate_Score) > 15)
                        {


                            if (list_pre_customers[i].customer_data[j].loan_details[0].Loan_Authority.Length >= 30)
                            {
                                list_pre_customers[i].customer_data[j].loan_details[0].Loan_Authority = "無法解析債權者姓名";
                            }
                            index += 1;
                            string loan_more_data = "";
                            for (int x = 0; x <= list_pre_customers[i].customer_data[j].loan_details.Count - 1; x++)
                            {
                                try
                                {


                                string[] area_arr = list_pre_customers[i].information_total_area.Split('-');
                                Excel_WS.Cells[index, 1] = index.ToString();
                                Excel_WS.Cells[index, 2] = list_pre_customers[i].customer_data[j].name;
                                Excel_WS.Cells[index, 3] = list_pre_customers[i].customer_data[j].gender;
                                Excel_WS.Cells[index, 4] = list_pre_customers[i].customer_data[j].idnetity;
                                Excel_WS.Cells[index, 5] = list_pre_customers[i].information_address;
                                Excel_WS.Cells[index, 6] = list_pre_customers[i].information_building_type;
                                Excel_WS.Cells[index, 7] = area_arr[0];
                                if (x == 0)
                                {
                                    Excel_WS.Cells[index, 8] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                    if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                    {
                                        Excel_WS.Cells[index, 9] = "無法解析此債權人姓名";
                                    }
                                    else
                                    {
                                        Excel_WS.Cells[index, 9] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                    }
                                    Excel_WS.Cells[index, 10] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                }
                                if (x == 1)
                                {
                                    Excel_WS.Cells[index, 11] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                    if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                    {
                                        Excel_WS.Cells[index, 12] = "無法解析此債權人姓名";
                                    }
                                    else
                                    {
                                        Excel_WS.Cells[index, 12] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                    }
                                    Excel_WS.Cells[index, 13] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                }
                                Excel_WS.Cells[index, 14] = "";
                                if (x >= 2)
                                {
                                    loan_more_data += "-----第" + (x + 1).ToString() + "胎----\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].register_date + "\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority + "\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount + "\r\n";
                                    Excel_WS.Cells[index, 15] = loan_more_data;
                                }
                                Excel_WS.Cells[index, 16] = list_pre_customers[i].customer_data[j].evaluate_Score;
                                Excel_WS.Cells[index, 17] = list_pre_customers[i].customer_data[j].droit_source;
                                Excel_WS.Cells[index, 18] = list_pre_customers[i].customer_data[j].evaluate_droit_source;
                                Excel_WS.Cells[index, 19] = area_arr[0];
                                Excel_WS.Cells[index, 20] = list_pre_customers[i].evaluate_total_area;
                                Excel_WS.Cells[index, 21] = list_pre_customers[i].customer_data[j].loan_tittle;
                                Excel_WS.Cells[index, 22] = list_pre_customers[i].customer_data[j].evaluate_loan_order_gap;
                                Excel_WS.Cells[index, 23] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                Excel_WS.Cells[index, 24] = list_pre_customers[i].customer_data[j].evaluate_rigester_Date;
                                Excel_WS.Cells[index, 25] = list_pre_customers[i].customer_data[j].holding_range;
                                Excel_WS.Cells[index, 26] = list_pre_customers[i].customer_data[j].evaluate_holder_area_part;
                                Excel_WS.Cells[index, 27] = list_pre_customers[i].information_total_floor;
                                Excel_WS.Cells[index, 28] = list_pre_customers[i].evaluate_total_floor;
                                Excel_WS.Cells[index, 29] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                Excel_WS.Cells[index, 30] = list_pre_customers[i].customer_data[j].evaluate_droit_rigester_first_name;
                                Excel_WS.Cells[index, 31] = "";
                                Excel_WS.Cells[index, 32] = "";
                                Excel_WS.Cells[index, 33] = "";
                                Excel_WS.Cells[index, 34] = "";
                                Excel_WS.Cells[index, 35] = list_pre_customers[i].information_building_section + "_" + list_pre_customers[i].information_building_buildingnumber;
                                Excel_WS.Cells[index, 36] = list_pre_customers[i].information_district;

                                }
                                catch (Exception)
                                {

                                    throw;
                                }
                            }
                        }
                    }
                }

                Excel_WS.Range["A1", "Z" + index].RowHeight = 35;
                Excel_WS = null;

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message.ToString());
            }
        }
        public void Sheet_Level_C(Excel.Application Excel_App, Excel.Workbook Excel_WB, List<Set_Pre_Customer_Data> list_pre_customers)
        {
            object missing = Type.Missing;
            Excel.Worksheet Excel_WS = Excel_WB.Sheets.Add(missing, missing, 1, missing) as Excel.Worksheet;
            Excel_WS.Name = "C級名單";
            Excel_WS.Cells.NumberFormat = "@";

            Excel_WS.Cells[1, 1] = "項次";
            Excel_WS.Cells[1, 2] = "客戶姓名";
            Excel_WS.Cells[1, 3] = "性別";
            Excel_WS.Cells[1, 4] = "身份證號";
            Excel_WS.Cells[1, 5] = "標的物地址 / 所有權地址標的物地址 / 所有權地址";
            Excel_WS.Cells[1, 6] = "建物類型";
            Excel_WS.Cells[1, 7] = "建物坪";
            Excel_WS.Cells[1, 8] = "一胎時間";
            Excel_WS.Cells[1, 9] = "一胎債權";
            Excel_WS.Cells[1, 10] = "一胎設定債權(萬)";
            Excel_WS.Cells[1, 11] = "二胎時間";
            Excel_WS.Cells[1, 12] = "二胎債權";
            Excel_WS.Cells[1, 13] = "二胎設定債權(萬)";
            Excel_WS.Cells[1, 14] = "備註";
            Excel_WS.Cells[1, 15] = "多胎以上資料（時間、債權、設定債權金額)";
            Excel_WS.Cells[1, 16] = "評分";
            Excel_WS.Cells[1, 17] = "登記原因";
            Excel_WS.Cells[1, 18] = "評分項目-登記原因";
            Excel_WS.Cells[1, 19] = "建物㎡";
            Excel_WS.Cells[1, 20] = "評分項目-建物㎡";
            Excel_WS.Cells[1, 21] = "他項登記次序";
            Excel_WS.Cells[1, 22] = "評分項目-他項登記次序";
            Excel_WS.Cells[1, 23] = "登記日期";
            Excel_WS.Cells[1, 24] = "評分項目-登記日期";
            Excel_WS.Cells[1, 25] = "權力範圍";
            Excel_WS.Cells[1, 26] = "評分項目-權力範圍";
            Excel_WS.Cells[1, 27] = "建物總樓層";
            Excel_WS.Cells[1, 28] = "評分項目-建物總樓層";
            Excel_WS.Cells[1, 29] = "權力人姓名";
            Excel_WS.Cells[1, 30] = "評分項目-權力人姓名";
            Excel_WS.Cells[1, 31] = "跑謄狀態";
            Excel_WS.Cells[1, 32] = "電話";
            Excel_WS.Cells[1, 33] = "聯徵";
            Excel_WS.Cells[1, 34] = "開發日期";
            Excel_WS.Cells[1, 35] = "段號_建號";

            try
            {
                int index = 1;
                Loadding_Bar.Create_Loadding_UI loading_bar = new Loading_Bar.Create_Loadding_UI();////////////
                loading_bar.Set_Maximum(list_pre_customers.Count - 1);//////////////////

                for (int i = 0; i <= list_pre_customers.Count - 1; i++)
                {
                    loading_bar.Update_Loadding_UI(i + 5);/////////////


                    for (int j = 0; j <= list_pre_customers[i].customer_data.Count - 1; j++)
                    {
                        if (15 >= int.Parse(list_pre_customers[i].customer_data[j].evaluate_Score) && int.Parse(list_pre_customers[i].customer_data[j].evaluate_Score) > 8)
                        {


                            if (list_pre_customers[i].customer_data[j].loan_details[0].Loan_Authority.Length >= 30)
                            {
                                list_pre_customers[i].customer_data[j].loan_details[0].Loan_Authority = "無法解析債權者姓名";
                            }
                            index += 1;
                            string loan_more_data = "";
                            for (int x = 0; x <= list_pre_customers[i].customer_data[j].loan_details.Count - 1; x++)
                            {
                                try
                                {

                                
                                string[] area_arr = list_pre_customers[i].information_total_area.Split('-');
                                Excel_WS.Cells[index, 1] = index.ToString();
                                Excel_WS.Cells[index, 2] = list_pre_customers[i].customer_data[j].name;
                                Excel_WS.Cells[index, 3] = list_pre_customers[i].customer_data[j].gender;
                                Excel_WS.Cells[index, 4] = list_pre_customers[i].customer_data[j].idnetity;
                                Excel_WS.Cells[index, 5] = list_pre_customers[i].information_address;
                                Excel_WS.Cells[index, 6] = list_pre_customers[i].information_building_type;
                                Excel_WS.Cells[index, 7] = area_arr[0];
                                if (x == 0)
                                {
                                    Excel_WS.Cells[index, 8] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                    if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                    {
                                        Excel_WS.Cells[index, 9] = "無法解析此債權人姓名";
                                    }
                                    else
                                    {
                                        Excel_WS.Cells[index, 9] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                    }
                                    Excel_WS.Cells[index, 10] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                }
                                if (x == 1)
                                {
                                    Excel_WS.Cells[index, 11] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                    if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                    {
                                        Excel_WS.Cells[index, 12] = "無法解析此債權人姓名";
                                    }
                                    else
                                    {
                                        Excel_WS.Cells[index, 12] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                    }
                                    Excel_WS.Cells[index, 13] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                }
                                Excel_WS.Cells[index, 14] = "";
                                if (x >= 2)
                                {
                                    loan_more_data += "-----第" + (x + 1).ToString() + "胎----\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].register_date + "\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority + "\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount + "\r\n";
                                    Excel_WS.Cells[index, 15] = loan_more_data;
                                }
                                Excel_WS.Cells[index, 16] = list_pre_customers[i].customer_data[j].evaluate_Score;
                                Excel_WS.Cells[index, 17] = list_pre_customers[i].customer_data[j].droit_source;
                                Excel_WS.Cells[index, 18] = list_pre_customers[i].customer_data[j].evaluate_droit_source;
                                Excel_WS.Cells[index, 19] = area_arr[0];
                                Excel_WS.Cells[index, 20] = list_pre_customers[i].evaluate_total_area;
                                Excel_WS.Cells[index, 21] = list_pre_customers[i].customer_data[j].loan_tittle;
                                Excel_WS.Cells[index, 22] = list_pre_customers[i].customer_data[j].evaluate_loan_order_gap;
                                Excel_WS.Cells[index, 23] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                Excel_WS.Cells[index, 24] = list_pre_customers[i].customer_data[j].evaluate_rigester_Date;
                                Excel_WS.Cells[index, 25] = list_pre_customers[i].customer_data[j].holding_range;
                                Excel_WS.Cells[index, 26] = list_pre_customers[i].customer_data[j].evaluate_holder_area_part;
                                Excel_WS.Cells[index, 27] = list_pre_customers[i].information_total_floor;
                                Excel_WS.Cells[index, 28] = list_pre_customers[i].evaluate_total_floor;
                                Excel_WS.Cells[index, 29] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                Excel_WS.Cells[index, 30] = list_pre_customers[i].customer_data[j].evaluate_droit_rigester_first_name;
                                Excel_WS.Cells[index, 31] = "";
                                Excel_WS.Cells[index, 32] = "";
                                Excel_WS.Cells[index, 33] = "";
                                Excel_WS.Cells[index, 34] = "";
                                Excel_WS.Cells[index, 35] = list_pre_customers[i].information_building_section + "_" + list_pre_customers[i].information_building_buildingnumber;
                                Excel_WS.Cells[index, 36] = list_pre_customers[i].information_district;
                                }
                                catch (Exception)
                                {

                                    throw;
                                }
                            }
                        }
                    }
                }

                Excel_WS.Range["A1", "Z" + index].RowHeight = 35;
                Excel_WS = null;

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message.ToString());
            }
        }
        public void Sheet_Level_D(Excel.Application Excel_App, Excel.Workbook Excel_WB, List<Set_Pre_Customer_Data> list_pre_customers)
        {
            object missing = Type.Missing;
            Excel.Worksheet Excel_WS = Excel_WB.Sheets.Add(missing, missing, 1, missing) as Excel.Worksheet;
            Excel_WS.Name = "D級名單";
            Excel_WS.Cells.NumberFormat = "@";

            Excel_WS.Cells[1, 1] = "項次";
            Excel_WS.Cells[1, 2] = "客戶姓名";
            Excel_WS.Cells[1, 3] = "性別";
            Excel_WS.Cells[1, 4] = "身份證號";
            Excel_WS.Cells[1, 5] = "標的物地址 / 所有權地址標的物地址 / 所有權地址";
            Excel_WS.Cells[1, 6] = "建物類型";
            Excel_WS.Cells[1, 7] = "建物坪";
            Excel_WS.Cells[1, 8] = "一胎時間";
            Excel_WS.Cells[1, 9] = "一胎債權";
            Excel_WS.Cells[1, 10] = "一胎設定債權(萬)";
            Excel_WS.Cells[1, 11] = "二胎時間";
            Excel_WS.Cells[1, 12] = "二胎債權";
            Excel_WS.Cells[1, 13] = "二胎設定債權(萬)";
            Excel_WS.Cells[1, 14] = "備註";
            Excel_WS.Cells[1, 15] = "多胎以上資料（時間、債權、設定債權金額)";
            Excel_WS.Cells[1, 16] = "評分";
            Excel_WS.Cells[1, 17] = "登記原因";
            Excel_WS.Cells[1, 18] = "評分項目-登記原因";
            Excel_WS.Cells[1, 19] = "建物㎡";
            Excel_WS.Cells[1, 20] = "評分項目-建物㎡";
            Excel_WS.Cells[1, 21] = "他項登記次序";
            Excel_WS.Cells[1, 22] = "評分項目-他項登記次序";
            Excel_WS.Cells[1, 23] = "登記日期";
            Excel_WS.Cells[1, 24] = "評分項目-登記日期";
            Excel_WS.Cells[1, 25] = "權力範圍";
            Excel_WS.Cells[1, 26] = "評分項目-權力範圍";
            Excel_WS.Cells[1, 27] = "建物總樓層";
            Excel_WS.Cells[1, 28] = "評分項目-建物總樓層";
            Excel_WS.Cells[1, 29] = "權力人姓名";
            Excel_WS.Cells[1, 30] = "評分項目-權力人姓名";
            Excel_WS.Cells[1, 31] = "跑謄狀態";
            Excel_WS.Cells[1, 32] = "電話";
            Excel_WS.Cells[1, 33] = "聯徵";
            Excel_WS.Cells[1, 34] = "開發日期";
            Excel_WS.Cells[1, 35] = "段號_建號";

            try
            {
                int index = 1; 
                Loadding_Bar.Create_Loadding_UI loading_bar = new Loading_Bar.Create_Loadding_UI();////////////
                loading_bar.Set_Maximum(list_pre_customers.Count - 1);//////////////////

                for (int i = 0; i <= list_pre_customers.Count - 1; i++)
                {
                    loading_bar.Update_Loadding_UI(i + 5);/////////////


                    for (int j = 0; j <= list_pre_customers[i].customer_data.Count - 1; j++)
                    {
                        if (7 >= int.Parse(list_pre_customers[i].customer_data[j].evaluate_Score))
                        {


                            if (list_pre_customers[i].customer_data[j].loan_details[0].Loan_Authority.Length >= 30)
                            {
                                list_pre_customers[i].customer_data[j].loan_details[0].Loan_Authority = "無法解析債權者姓名";
                            }
                            index += 1;
                            string loan_more_data = "";
                            for (int x = 0; x <= list_pre_customers[i].customer_data[j].loan_details.Count - 1; x++)
                            {
                                try
                                {


                                string[] area_arr = list_pre_customers[i].information_total_area.Split('-');
                                Excel_WS.Cells[index, 1] = index.ToString();
                                Excel_WS.Cells[index, 2] = list_pre_customers[i].customer_data[j].name;
                                Excel_WS.Cells[index, 3] = list_pre_customers[i].customer_data[j].gender;
                                Excel_WS.Cells[index, 4] = list_pre_customers[i].customer_data[j].idnetity;
                                Excel_WS.Cells[index, 5] = list_pre_customers[i].information_address;
                                Excel_WS.Cells[index, 6] = list_pre_customers[i].information_building_type;
                                Excel_WS.Cells[index, 7] = area_arr[0];
                                if (x == 0)
                                {
                                    Excel_WS.Cells[index, 8] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                    if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                    {
                                        Excel_WS.Cells[index, 9] = "無法解析此債權人姓名";
                                    }
                                    else
                                    {
                                        Excel_WS.Cells[index, 9] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                    }
                                    Excel_WS.Cells[index, 10] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                }
                                if (x == 1)
                                {
                                    Excel_WS.Cells[index, 11] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                    if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Length >= 100)
                                    {
                                        Excel_WS.Cells[index, 12] = "無法解析此債權人姓名";
                                    }
                                    else
                                    {
                                        Excel_WS.Cells[index, 12] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                    }
                                    Excel_WS.Cells[index, 13] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount;
                                }
                                Excel_WS.Cells[index, 14] = "";
                                if (x >= 2)
                                {
                                    loan_more_data += "-----第" + (x + 1).ToString() + "胎----\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].register_date + "\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority + "\r\n";
                                    loan_more_data += list_pre_customers[i].customer_data[j].loan_details[x].Loan_Amount + "\r\n";
                                    Excel_WS.Cells[index, 15] = loan_more_data;
                                }
                                Excel_WS.Cells[index, 16] = list_pre_customers[i].customer_data[j].evaluate_Score;
                                Excel_WS.Cells[index, 17] = list_pre_customers[i].customer_data[j].droit_source;
                                Excel_WS.Cells[index, 18] = list_pre_customers[i].customer_data[j].evaluate_droit_source;
                                Excel_WS.Cells[index, 19] = area_arr[0];
                                Excel_WS.Cells[index, 20] = list_pre_customers[i].evaluate_total_area;
                                Excel_WS.Cells[index, 21] = list_pre_customers[i].customer_data[j].loan_tittle;
                                Excel_WS.Cells[index, 22] = list_pre_customers[i].customer_data[j].evaluate_loan_order_gap;
                                Excel_WS.Cells[index, 23] = list_pre_customers[i].customer_data[j].loan_details[x].register_date;
                                Excel_WS.Cells[index, 24] = list_pre_customers[i].customer_data[j].evaluate_rigester_Date;
                                Excel_WS.Cells[index, 25] = list_pre_customers[i].customer_data[j].holding_range;
                                Excel_WS.Cells[index, 26] = list_pre_customers[i].customer_data[j].evaluate_holder_area_part;
                                Excel_WS.Cells[index, 27] = list_pre_customers[i].information_total_floor;
                                Excel_WS.Cells[index, 28] = list_pre_customers[i].evaluate_total_floor;
                                Excel_WS.Cells[index, 29] = list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority;
                                Excel_WS.Cells[index, 30] = list_pre_customers[i].customer_data[j].evaluate_droit_rigester_first_name;
                                Excel_WS.Cells[index, 31] = "";
                                Excel_WS.Cells[index, 32] = "";
                                Excel_WS.Cells[index, 33] = "";
                                Excel_WS.Cells[index, 34] = "";
                                Excel_WS.Cells[index, 35] = list_pre_customers[i].information_building_section + "_" + list_pre_customers[i].information_building_buildingnumber;
                                Excel_WS.Cells[index, 36] = list_pre_customers[i].information_district;
                                }
                                catch (Exception)
                                {

                                    throw;
                                }
                            }
                        }
                    }
                }

                Excel_WS.Range["A1", "Z" + index].RowHeight = 35;
                Excel_WS = null;

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message.ToString());
            }
        }

        public void Sheet_Level_A_New(Excel.Workbook Excel_WB,Excel.Worksheet Reference_Sheet)
        {
            Loadding_Bar.Create_Loadding_UI loading_bar = new Loading_Bar.Create_Loadding_UI();////////////

            object missing = Type.Missing;
            Excel.Worksheet Excel_WS = Excel_WB.Sheets.Add(missing, missing, 1, missing) as Excel.Worksheet;
            Excel_WS.Name = "Level_A";
            Set_Font_Style(ref Excel_WS);
            Set_Tittle_Freeze(ref Excel_WS);

            int Use_Rows_Int = Reference_Sheet.UsedRange.Cells.Rows.Count; //得到行数

            Range Reference_Sheet_Data_Range = Reference_Sheet.Cells.get_Range("A1", "AJ" + Use_Rows_Int);  //讀取資料
            object[,] Data_Arr = (object[,])Reference_Sheet_Data_Range.Value2;   //建立二維陣列
            int Range = (int)Math.Round((Reference_Sheet_Data_Range.Rows.Count-1) * 0.25, 0); // 區間值

            loading_bar.Set_Maximum(Range);

            int index = 1;

            for (int i = 1; i <= Range+1; i++)
            {
                for (int j = 1; j <= Reference_Sheet_Data_Range.Columns.Count; j++)
                {
                    if (j == 25)
                    {
                        Excel_WS.Cells[index, 25].NumberFormat = "@";
                    }
                    Excel_WS.Cells[index, j] = Data_Arr[i,j];


                }
                string range_str = string.Format("A{0}:AJ{1}", 1.ToString(), 1.ToString());
                Set_Tittle_Column_Style(ref Excel_WS, range_str);

                range_str = string.Format("A{0}:AJ{1}", index, index);
                Set_Column_Style(ref Excel_WS, range_str);


                index++;

                loading_bar.Update_Loadding_UI(index);

            }
            Set_Column_Currency_NumberFormat(ref Excel_WS, index);
            Excel_WS.Range["A1", "AJ" + index].RowHeight = 35;



        }
        public void Sheet_Level_B_New(Excel.Workbook Excel_WB, Excel.Worksheet Reference_Sheet)
        {
            Loadding_Bar.Create_Loadding_UI loading_bar = new Loading_Bar.Create_Loadding_UI();////////////


            object missing = Type.Missing;
            Excel.Worksheet Excel_WS = Excel_WB.Sheets.Add(missing, missing, 1, missing) as Excel.Worksheet;
            Excel_WS.Name = "Level_B";
            Set_Font_Style(ref Excel_WS);
            Set_Tittle_Freeze(ref Excel_WS);

            int Use_Rows_Int = Reference_Sheet.UsedRange.Cells.Rows.Count; //得到行数
            Range Reference_Sheet_Data_Range = Reference_Sheet.Cells.get_Range("A1", "AJ" + Use_Rows_Int);  //讀取資料
            object[,] Data_Arr = (object[,])Reference_Sheet_Data_Range.Value2;   //建立二維陣列

            int Range = (int)Math.Round((Reference_Sheet_Data_Range.Rows.Count - 1) * 0.25, 0);
            int index = 2;
            loading_bar.Set_Maximum((Range * 2 + 1)-(Range + 2));

            //填入標題欄
            for (int i = 1; i <= Reference_Sheet_Data_Range.Columns.Count ; i++)
            {
                Excel_WS.Cells[1, i] = Data_Arr[1, i];
            }



            for (int i = Range+2; i <= Range*2+1; i++)
            {
                for (int j = 1; j <= Reference_Sheet_Data_Range.Columns.Count; j++)
                {
                    if (j == 25)
                    {
                        Excel_WS.Cells[index, 25].NumberFormat = "@";
                    }
                    Excel_WS.Cells[index, j] = Data_Arr[i, j];
                }
                string range_str = string.Format("A{0}:AJ{1}", 1.ToString(), 1.ToString());
                Set_Tittle_Column_Style(ref Excel_WS, range_str);

                range_str = string.Format("A{0}:AJ{1}", index, index);
                Set_Column_Style(ref Excel_WS, range_str);



                index++;
                loading_bar.Update_Loadding_UI(index);
            }

            Set_Column_Currency_NumberFormat(ref Excel_WS, index);
            Excel_WS.Range["A1", "AJ" + index].RowHeight = 35;


        }
        public void Sheet_Level_C_New(Excel.Workbook Excel_WB, Excel.Worksheet Reference_Sheet)
        {
            Loadding_Bar.Create_Loadding_UI loading_bar = new Loading_Bar.Create_Loadding_UI();////////////

            object missing = Type.Missing;
            Excel.Worksheet Excel_WS = Excel_WB.Sheets.Add(missing, missing, 1, missing) as Excel.Worksheet;
            Excel_WS.Name = "Level_C";
            Set_Font_Style(ref Excel_WS);
            Set_Tittle_Freeze(ref Excel_WS);

            int Use_Rows_Int = Reference_Sheet.UsedRange.Cells.Rows.Count; //得到行数
            Range Reference_Sheet_Data_Range = Reference_Sheet.Cells.get_Range("A1", "AJ" + Use_Rows_Int);  //讀取資料
            object[,] Data_Arr = (object[,])Reference_Sheet_Data_Range.Value2;   //建立二維陣列

            int Range = (int)Math.Round((Reference_Sheet_Data_Range.Rows.Count - 1) * 0.25, 0);
            int index = 2;


            loading_bar.Set_Maximum(Reference_Sheet_Data_Range.Rows.Count-(Range * 2 +2)) ;

            //填入標題欄
            for (int i = 1; i <= Reference_Sheet_Data_Range.Columns.Count; i++)
            {
                Excel_WS.Cells[1, i] = Data_Arr[1, i];
            }


            for (int i = (Range*2)+2; i <= Reference_Sheet_Data_Range.Rows.Count; i++)
            {
                for (int j = 1; j <= Reference_Sheet_Data_Range.Columns.Count; j++)
                {
                    if (j == 25)
                    {
                        Excel_WS.Cells[index, 25].NumberFormat = "@";
                    }
                    Excel_WS.Cells[index, j] = Data_Arr[i, j];
                }
                string range_str = string.Format("A{0}:AJ{1}", 1.ToString(), 1.ToString());
                Set_Tittle_Column_Style(ref Excel_WS, range_str);

                range_str = string.Format("A{0}:AJ{1}", index, index);
                Set_Column_Style(ref Excel_WS, range_str);


                index++;
                loading_bar.Update_Loadding_UI(index);
            }

            Set_Column_Currency_NumberFormat(ref Excel_WS, index);
            Excel_WS.Range["A1", "AJ" + index].RowHeight = 35;

        }

        public void Set_Column_Style(ref Excel.Worksheet Excel_WS,string range_str)
        {
            Range data_range = Excel_WS.Range[range_str];
            data_range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            data_range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            data_range.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            data_range.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        }
        public void Set_Tittle_Column_Style(ref Excel.Worksheet Excel_WS, string range_str)
        {
            Range data_range = Excel_WS.Range[range_str];
            data_range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            data_range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            data_range.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
            data_range.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
        }

        public void Set_Tittle_Freeze(ref Excel.Worksheet Excel_WS)
        {
            Excel_WS.Activate();
            Excel_WS.Application.ActiveWindow.SplitRow = 1;
            Excel_WS.Application.ActiveWindow.FreezePanes = true;
        }
        public void Set_Font_Style(ref Excel.Worksheet Excel_WS)
        {
            Excel_WS.Cells.Font.Size = 9;
            Excel_WS.Cells.Font.Name = "微軟正黑體";
            Excel_WS.Cells.Font.FontStyle = "bold";

        }
        public void Set_Column_Currency_NumberFormat(ref Excel.Worksheet Excel_WS,int index)
        {
            Excel_WS.Range["J2","J" + index.ToString()].NumberFormat = "#,###";
        }

        public void Recording_Error_Data(Excel.Worksheet Excel_Error_Sheet,ref int Error_index,string Current_BuildBaseID_CustomerIdentity)
        {
            Error_index += 1;// 記錄錯誤用
            string[] Error_Arr = Current_BuildBaseID_CustomerIdentity.Split('_');
            Excel_Error_Sheet.Cells[Error_index, 1] = Error_Arr[0];// BuildBaseID
            Excel_Error_Sheet.Cells[Error_index, 2] = Error_Arr[1];// Cutomer_Identity
        }



        public void Sort(ref Excel.Worksheet Excel_WS,string Range_Index,string filter_data)
        {
            /// 資料格式：評分、區域
            string[] filter_data_arr = filter_data.Split(',');



            Excel_WS.Range[filter_data_arr[0] + "1", filter_data_arr[0] + Range_Index].NumberFormat = "General";
            Excel_WS.Range[filter_data_arr[0] + "1", filter_data_arr[0] + Range_Index].Interior.Color = Color.Yellow;
            // Set sort properties            
            Excel_WS.Sort.SetRange(Excel_WS.Range["A1", filter_data_arr[1] + Range_Index]);
            Excel_WS.Sort.Header = Excel.XlYesNoGuess.xlYes;
            if (filter_data_arr[2].Equals("District"))
            {
                Excel_WS.Sort.SortFields.Add(Excel_WS.Range[filter_data_arr[1] + "1", filter_data_arr[1] + Range_Index], Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending);
            }
            else if (filter_data_arr[2].Equals("Score"))
            {
                Excel_WS.Sort.SortFields.Add(Excel_WS.Range[filter_data_arr[0] + "1", filter_data_arr[0] + Range_Index], Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending);
            }

            // Sort worksheet
            Excel_WS.Sort.Apply();


        }


    }
}
