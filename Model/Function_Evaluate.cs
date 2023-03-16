using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CADVISER_Transcript_Customers.Model.Data_Set;
namespace CADVISER_Transcript_Customers.Model.Data_Set
{
    public class Function_Evaluate
    {

        public List<Set_Pre_Customer_Data> Register_Date(List<Set_Pre_Customer_Data> list_pre_customers)
        {

            for (int i = 0; i <= list_pre_customers.Count - 1; i++)
            {
                for (int j = 0; j < list_pre_customers[i].customer_data.Count; j++)
                {
                    int Score = 0;
                    int Temp_Score = 0;

                    if(list_pre_customers[i].customer_data[j].loan_details.Count != null) { 
                        for (int z = 0; z <= list_pre_customers[i].customer_data[j].loan_details.Count-1 ; z++)
                        {
                            DateTime Current_Time = DateTime.Now;
                            DateTime Register_Time = Convert.ToDateTime(list_pre_customers[i].customer_data[j].loan_details[z].register_date);
                            TimeSpan diff = Current_Time.ToUniversalTime() - Register_Time.ToUniversalTime();

                            double diff_years = diff.TotalDays / 365;

                            if ( diff_years < 10)
                            {
                                Score = 1;
                            }
                            if( diff_years >= 10)
                            {
                                Score = 2;
                            }
                            if ( diff_years >= 15)
                            {
                                Score = 3;
                            }
                            if (diff_years >= 20)
                            {
                                Score = 4;
                            }
                            if ((Temp_Score == 0) || (Temp_Score > Score)){
                                Temp_Score = Score;
                            }
                        }
                    }
                    list_pre_customers[i].customer_data[j].evaluate_rigester_Date = Temp_Score.ToString();
                }
            }
            return list_pre_customers;
        }
        public List<Set_Pre_Customer_Data> Loan_Order_Gap(List<Set_Pre_Customer_Data> list_pre_customers)
        {
            for (int i = 0; i <= list_pre_customers.Count-1 ; i++)
            {
                for (int j = 0; j < list_pre_customers[i].customer_data.Count; j++)
                {
                    int dealta = 0;
                    string[] loan_order_arr = list_pre_customers[i].customer_data[j].loan_tittle.ToString().Split('｜');

                    if(loan_order_arr.Length == 1)
                    {
                        list_pre_customers[i].customer_data[j].evaluate_loan_order_gap = (1).ToString();
                    }
                    else
                    {
                        int Max = 0;
                        int Min = 0;

                        for (int x = 0; x <= loan_order_arr.Length-1; x++)
                        {
                            string[] compare_loan_order_arr = loan_order_arr[x].Split('-');
                            if (int.Parse(compare_loan_order_arr[0]) > Max)
                            {
                                Max = int.Parse(compare_loan_order_arr[0]);
                            }
                            if((Min == 0 ) || int.Parse(compare_loan_order_arr[0]) < Min)
                            {
                                Min = int.Parse(compare_loan_order_arr[0]);
                            }
                        }
                        int Score = 0;
                        int loan_order_gap = (Max - Min);
                        if(loan_order_gap >= 8)
                        {
                            Score = 4;
                        }
                        if (loan_order_gap >= 4 && loan_order_gap <= 7)
                        {
                            Score = 3;
                        }
                        if (loan_order_gap >= 2 && loan_order_gap <= 3)
                        {
                            Score = 2;
                        }
                        if (loan_order_gap == 1)
                        {
                            Score = 1;
                        }

                        list_pre_customers[i].customer_data[j].evaluate_loan_order_gap = Score.ToString();
                    }


                }
            }



            return list_pre_customers;
        }
        public List<Set_Pre_Customer_Data> Holder_Area_Part(List<Set_Pre_Customer_Data> list_pre_customers)
        {

            for (int i = 0; i <= list_pre_customers.Count -1  ; i++)
            {
                int Score = 0;
                for (int j = 0; j <= list_pre_customers[i].customer_data.Count-1 ; j++)
                {
                    if(list_pre_customers[i].customer_data[j].holding_range.Contains("公同共有")) { list_pre_customers[i].customer_data[j].holding_range = "1000/0"; }
                    string[] holding_range_arr = list_pre_customers[i].customer_data[j].holding_range.Split('/');
                    double holding_value = double.Parse(holding_range_arr[1]) / double.Parse(holding_range_arr[0]);

                    if(holding_value == 1)
                    {
                        Score = 4;
                    }
                    if(holding_value >= 0.5 &&　holding_value <= 0.99)
                    {
                        Score = 3;
                    }
                    if(holding_value < 0.5)
                    {
                        Score = 1;
                    }
                    if(holding_value == 0)
                    {
                        Score = 0;
                    }

                    list_pre_customers[i].customer_data[j].evaluate_holder_area_part = Score.ToString();

                }
            }



            return list_pre_customers;
        }
        public List<Set_Pre_Customer_Data> Register_Reason(List<Set_Pre_Customer_Data> list_pre_customers)
        {
            /// 開發初期不選擇走 二微陣列，
            string[] Score_Four = { "繼承", "剩餘財產差額分配", "預告登記", "夫妻贈與", "遺囑繼承", "塗銷信託", "塗銷查封", "判決繼承", "遺贈", "贈與", "調解回復所有權", "和解共有物分割" };
            string[] Score_Three = { "信託", "受託人變更", "拋棄", "分割繼承", "調解繼承", "判決移轉", "判決塗銷" };
            string[] Score_Two = { "調解移轉", "第一次登記", "判決共有物分割", "和解移轉", "拍賣", "土地重劃" };
            string[] Score_One = { "判決回復所有權", "調解共有物分割", "買賣", "合併", "註記", "共有物分割", "交換" };
            string[] Score_Zero = { "法人合併" };
            int Score = 0;


            for (int i = 0; i <= list_pre_customers.Count-1 ; i++)
            {
                for (int j = 0; j <= list_pre_customers[i].customer_data.Count -1 ; j++)
                {
                    for (int x = 0; x <= Score_Four.Length-1 ; x++)
                    {
                        if (list_pre_customers[i].customer_data[j].droit_source.Equals(Score_Four[x]))
                        {
                            Score = 4;
                        }
                    }
                    for (int x = 0; x <= Score_Three.Length - 1; x++)
                    {
                        if (list_pre_customers[i].customer_data[j].droit_source.Equals(Score_Three[x]))
                        {
                            Score = 3;
                        }
                    }
                    for (int x = 0; x <= Score_Two.Length - 1; x++)
                    {
                        if (list_pre_customers[i].customer_data[j].droit_source.Equals(Score_Two[x]))
                        {
                            Score = 2;
                        }
                    }
                    for (int x = 0; x <= Score_One.Length - 1; x++)
                    {
                        if (list_pre_customers[i].customer_data[j].droit_source.Equals(Score_One[x]))
                        {
                            Score = 1;
                        }
                    }
                    for (int x = 0; x <= Score_Zero.Length - 1; x++)
                    {
                        if (list_pre_customers[i].customer_data[j].droit_source.Equals(Score_Zero[x]))
                        {
                            Score = 0;
                        }
                    }
                    list_pre_customers[i].customer_data[j].evaluate_droit_source = Score.ToString();
                }

            }



            return list_pre_customers;
        }
        public List<Set_Pre_Customer_Data> Total_Area(List<Set_Pre_Customer_Data> list_pre_customers)
        {
            for (int i = 0; i <= list_pre_customers.Count-1 ; i++)
            {
                string[] Area_arr = list_pre_customers[i].information_total_area.Split('-');

                int Score = 0;
                if (!list_pre_customers[i].information_total_area.Equals("") && list_pre_customers[i].information_total_area.Length <= 100)
                {
                    if (double.Parse(Area_arr[0]) > 49)
                    {
                        Score = 4;
                    }
                    if (double.Parse(Area_arr[0]) <= 49)
                    {
                        Score = 1;
                    }
                }


                list_pre_customers[i].evaluate_total_area = Score.ToString();
            }
            return list_pre_customers;
        }
        public List<Set_Pre_Customer_Data> Total_Floor(List<Set_Pre_Customer_Data> list_pre_customers)
        {
            for (int i = 0; i <= list_pre_customers.Count-1 ; i++)
            {
                int Score = 0;
                if (!list_pre_customers[i].information_total_floor.Equals("")) {  
                    if (5 >= (int.Parse(list_pre_customers[i].information_total_floor)))
                    {
                        Score = 4;
                    }
                    if ((10 >= int.Parse(list_pre_customers[i].information_total_floor)) && (int.Parse(list_pre_customers[i].information_total_floor) > 5))
                    {
                        Score = 3;
                    }
                    if ((20 > int.Parse(list_pre_customers[i].information_total_floor)) && (int.Parse(list_pre_customers[i].information_total_floor) > 10))
                    {
                        Score = 2;
                    }
                    if (20 <= (int.Parse(list_pre_customers[i].information_total_floor)))
                    {
                        Score = 1;
                    }
                }
                list_pre_customers[i].evaluate_total_floor = Score.ToString();
            }
            return list_pre_customers;
        }
        public List<Set_Pre_Customer_Data> Sum_Total_Score(List<Set_Pre_Customer_Data> list_pre_customers)
        {

            for (int i = 0; i <= list_pre_customers.Count-1 ; i++)
            {
                int Total_Score = 0;
                //Total_Score += int.Parse(list_pre_customers[i].evaluate_droit_rigester_first_name);

                for (int j = 0; j <= list_pre_customers[i].customer_data.Count-1; j++)
                {
                    Total_Score += int.Parse(list_pre_customers[i].evaluate_total_area);
                    Total_Score += int.Parse(list_pre_customers[i].evaluate_total_floor);
                    Total_Score += int.Parse(list_pre_customers[i].customer_data[j].evaluate_droit_source);
                    Total_Score += int.Parse(list_pre_customers[i].customer_data[j].evaluate_holder_area_part);
                    Total_Score += int.Parse(list_pre_customers[i].customer_data[j].evaluate_loan_order_gap);
                    Total_Score += int.Parse(list_pre_customers[i].customer_data[j].evaluate_rigester_Date);
                    Total_Score += int.Parse(list_pre_customers[i].customer_data[j].evaluate_droit_rigester_first_name);
                    list_pre_customers[i].customer_data[j].evaluate_Score = Total_Score.ToString();
                    Total_Score = 0;
                }
            }
            return list_pre_customers;
        }
        public List<Set_Pre_Customer_Data> Evaluate_Droit_Rigester_First_Name(List<Set_Pre_Customer_Data> list_pre_customers)
        {
            for (int i = 0; i <= list_pre_customers.Count-1 ; i++)
            {
                for (int j = 0; j <= list_pre_customers[i].customer_data.Count-1 ; j++)
                {
                    bool Score_Four_Compare = false;
                    bool Zero_Four_Compare = false;

                    for (int x = 0; x <= list_pre_customers[i].customer_data[j].loan_details.Count - 1; x++)
                    {
                        string[] Droit_Register_First_Name_Arr = { "紀", "韓", "資融", "投資", "當鋪" };

                        for (int z = 0; z <= Droit_Register_First_Name_Arr.Length-4; z++)
                        {
                            if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Trim().Substring(0, 1).Equals(Droit_Register_First_Name_Arr[z]))
                            {
                                Score_Four_Compare=true;
                            }
                            if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Trim().Contains("融鎰"))
                            {
                                Zero_Four_Compare=true;
                            }
                        }

                        for (int z = 2; z <= Droit_Register_First_Name_Arr.Length - 1; z++)
                        {
                            if (list_pre_customers[i].customer_data[j].loan_details[x].Loan_Authority.Trim().Contains(Droit_Register_First_Name_Arr[z]))
                            {
                                Score_Four_Compare = true;
                            }
                        }

                    }
                    if (Score_Four_Compare)
                    {
                        list_pre_customers[i].customer_data[j].evaluate_droit_rigester_first_name = 4.ToString();
                    }
                    else if(Zero_Four_Compare)
                    {
                        list_pre_customers[i].customer_data[j].evaluate_droit_rigester_first_name = 0.ToString();
                    }
                    else
                    {
                        list_pre_customers[i].customer_data[j].evaluate_droit_rigester_first_name = 1.ToString();
                    }
                }
            }


            return list_pre_customers;
        }
    }
}
