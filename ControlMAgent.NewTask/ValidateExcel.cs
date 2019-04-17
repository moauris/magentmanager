using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EXCEL = Microsoft.Office.Interop.Excel;
using ControlMAgent.Base;
using System.Text.RegularExpressions;

namespace ControlMAgent.NewTask
{
    public class ValidateExcel
    {
        public bool IsValid { get; }
        public string InvalidMessage { get; }

        public ValidateExcel(EXCEL.Worksheet xlSheet)
        {
            IsValid = false; InvalidMessage = "";
            Regex rex = new Regex(@"^M/Agent.*");
            //Match M/Agent導入申請書(JP) or M/Agent Request(EN)

            bool Check_D5 = xlSheet.Range["$D$5"].Value == null ?
                false : rex.IsMatch(xlSheet.Range["$D$5"].Value);

            rex = new Regex(@"^\w{3}(\d|\w){5}(\.|$)");
            //Match a hostname 3 word, 5 word / Num hyprid, finish with either a . or end
            //Valid = true, not valid = false
            bool Check_H51 = xlSheet.Range["$H$51"].Value == null ?
                false : rex.IsMatch(xlSheet.Range["$H$51"].Value);
            bool Check_H84 = !(xlSheet.Range["$H$84"].Value == null);
            bool Check_H98 = !(xlSheet.Range["$H$98"].Value == null);

            StringBuilder sbMessage = new StringBuilder();
            if (!Check_D5) sbMessage.AppendLine("M/Agent Header not Found At $D$5");
            if (!Check_H51) sbMessage.AppendLine("Column 1 Primary Hostname is not Valid Format");
            if (!Check_H84) sbMessage.AppendLine("M/Agent Version not Specified.");
            if (!Check_H98) sbMessage.AppendLine("M/Server not Assigned." + xlSheet.Range["$H$98"].Value);
            InvalidMessage = sbMessage.ToString();
            if (InvalidMessage.Length == 0) IsValid = true;
        }


    }
}
