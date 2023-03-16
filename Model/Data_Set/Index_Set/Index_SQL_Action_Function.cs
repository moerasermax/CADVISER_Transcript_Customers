using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CADVISER_Transcript_Customers.Model.Data_Set.Index_Set
{
    public enum Index_SQL_Action_Function
    {
        /// --------測試--------
        Try_Connection = 1,
        /// --------新增--------
        Insert_Transcript_Customer_Data,
        /// --------獲取---------
        Get_Building,
        Get_Other,
        Get_Owner,
        Get_City,
        Get_Loan_Detail,
        Get_Building_Type

    }
}
