using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace ControlMAgent.CCM
{
    //This procedure query CCM.xls like a database, and stores results.
    public class CCMReader
    {
        private const string CCMExportDir = 
            @"C:\Users\MoChen\Downloads\CCM_Export\20190417_110100_CCM_Export.xls";
        

        public static DataTable FillDataset()
        {
            DataSet output = new DataSet();
            string ConnStr = string.Format(
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"{0}\";Extended Properties=\"Excel 8.0; HDR=Yes;IMEX=1\";"
            , CCMExportDir);
            using (OleDbDataAdapter adapter = new OleDbDataAdapter
                ("SELECT * FROM [20190417_110100_CCM_Export$];", ConnStr))
            {
                adapter.Fill(output, "CCM_Export");
                var outputTables = output.Tables["CCM_Export"];
                return outputTables;
            }

            
        }


    }
}
