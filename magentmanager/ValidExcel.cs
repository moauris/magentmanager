using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace magentmanager
{
    class ValidExcel
    {
        //Is the Form a Valid M/Agent Request?
        //Get Only
        public bool IsValid
        {
            get;
        }
        public string InvalidMessage
        {
            get;
        }

        //Constructor Initialize a Check on if the Form is Valid or not.
        //Then, populate the IsValid Field with true for Valid and False for InValid.
        //Critiria:
        // 1. D5 = "M/Agent導入申請書"
        // 2. H51 != {hostname}
        // 3. H81 != null (M/A Version)
        // 4. H98 != null (M/S Pri)
        public ValidExcel(EXCEL.Worksheet xlSheet)
        {
            IsValid = false; InvalidMessage = "";
            Regex rex = new Regex(@"^M/Agent.*");
            //Match M/Agent導入申請書(JP) or M/Agent Request(EN)
            
            bool Check_D5 = xlSheet.Range["$D$5"].Value == null?
                false : rex.IsMatch(xlSheet.Range["$D$5"].Value);

            rex = new Regex(@"^\w{3}(\d|\w){5}(\.|$)");
            //Match a hostname 3 word, 5 word / Num hyprid, finish with either a . or end
            //Valid = true, not valid = false
            bool Check_H51 = xlSheet.Range["$H$51"].Value == null?
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
