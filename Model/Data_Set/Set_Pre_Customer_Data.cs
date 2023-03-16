using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CADVISER_Transcript_Customers.Model.Data_Set
{
    public class Set_Pre_Customer_Data
    {
        public string buildingbaseId { get; set; }
        public List<Information_Customer> customer_data { get; set; }
        public string information_total_area { get; set; }
        public string information_total_floor { get; set; }
        public string information_address { get; set; }
        public string information_building_type { get; set; }
        public string information_building_section { get; set; }
        public string information_building_buildingnumber { get; set; }
        public string information_district { get; set; }
        public string evaluate_total_area { get; set; }
        public string evaluate_total_floor { get; set; }
        public string updating_time { get; set; }
    }


    public class Information_Customer
    {
        public string name { get; set; }
        public string idnetity { get; set; }
        public string gender { get; set; }
        public string droit_source { get; set; }
        public string holding_range { get; set; }
        public string loan_tittle { get; set; }
        public string evaluate_rigester_Date { get; set; }
        public string evaluate_loan_order_gap { get; set; }
        public string evaluate_holder_area_part { get; set; }
        public string evaluate_droit_source { get; set; }
        public string evaluate_droit_rigester_first_name { get; set; }
        public string evaluate_Score { get; set; }




        public List<Information__Loan_detail> loan_details { get; set; }
    }

    public class Information__Loan_detail
    {
        public string order_no { get; set; }
        public string register_date { get; set; }
        public string Loan_Authority { get; set; }
        public string Loan_Amount { get; set; }
    }
}
