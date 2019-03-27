using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.IO;
using EXCEL = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections.Generic;
using System.Reflection;
using System.Data.OleDb;
using System.Data;
using System.Collections;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Xml;
using ControlMAgent.Base;

namespace magentmanager
{
    class NewTask
    {
        const string SQLFolder =
            @"C:\Users\MoChen\source\repos\magentmanager\magentmanager\database\";
        const string DatabaseSource = SQLFolder + "magentr.accdb";
        const string DatabaseProvider = "Microsoft.ACE.OleDb.12.0";
        private static OleDbConnection DatabaseConnection()
        {
            var DbStringBuild = new OleDbConnectionStringBuilder
            {
                Provider = DatabaseProvider,
                DataSource = DatabaseSource
            };
            return new OleDbConnection(DbStringBuild.ToString());
        }

        // Invoked to Allow user to select file(s) to sync.
        internal static Task<FileInfo[]> TaskNewFileOpen()
        {
            //throw new NotImplementedException();
            List<FileInfo> listFiles = new List<FileInfo>();
            
            OpenFileDialog openFile = new OpenFileDialog
            {
                Title = "同期したいな M/Agent 申請書を選択してください",
                Multiselect = true, CheckFileExists = true,
                DefaultExt = ".xlsx;.xls",
                Filter = "Excel Worksheet (.xls;.xlsx)|*.xls;*xlsx"
            };
            if ((bool)openFile.ShowDialog())
            {
                //Clicked Yes
                foreach (string fs in openFile.FileNames)
                {
                    listFiles.Add(new FileInfo(fs));
                }
            }
            else
            {
                //Clicked No
                //throw new FileNotFoundException("No File Selected.");
                listFiles.Add(new FileInfo("null"));
            }
            return Task.FromResult(listFiles.ToArray());
        }

        //Sync Excel File to a XML.
        internal static async void TaskExcelToDatabase
            (FileInfo[] xlFiles, IProgress<bool> OnProgressChanged)
        {
            //throw new NotImplementedException();
            await Task.Run(() => OnProgressChanged.Report(true));

            IEnumerable<Process> ieProcessExcelBefore =
                from p in Process.GetProcessesByName("Excel")
                select p;

            EXCEL.Application xlApp = new EXCEL.Application();
            EXCEL.Workbooks xlWorkbooks = xlApp.Workbooks;
            IEnumerable<Process> ieProcessExcelNew =
                from p in Process.GetProcessesByName("Excel")
                where !ieProcessExcelBefore.Contains(p)
                select p;
            Process CurrentExcelProcess = ieProcessExcelNew.First();
            Debug.Print("New Excel Proc ID: "
                + CurrentExcelProcess.Id + " | "
                + CurrentExcelProcess.ProcessName);

            foreach (FileInfo f in xlFiles)
            {
                //Before Opening, Check if same name has existed in Database.
                bool ExcelIsNew = CheckDatabaseExcelExist(f.Name);
                if (ExcelIsNew)
                {
                    EXCEL.Workbook xlWbk = xlWorkbooks.Open(f.FullName); //Opening Excel File
                    EXCEL.Worksheet xlSheet = xlWbk.ActiveSheet;
                    //Before Sync Starts, Check if Worksheet is a valid Request.
                    ValidExcel validExcel = new ValidExcel(xlSheet);
                    if (validExcel.IsValid)
                    {
                        try
                        {
                            await ExcelToXML(f, xlSheet);
                        }
                        catch (Exception ex)
                        {
                            Debug.Print(ex.Message);
                            Debug.Print(ex.StackTrace);
                        }
                    }
                    else
                    {
                        Debug.Print(validExcel.InvalidMessage);
                    }

                    xlWbk.Close(false, Missing.Value, Missing.Value);

                }
                else
                {
                    Debug.Print("File Already Existed: " + f.Name);
                }
            }
            xlWorkbooks.Close();
            xlApp.Quit();
            GC.Collect();
            Console.Beep(1250, 20); Console.Beep(1650, 75);
            await Task.Run(() => CurrentExcelProcess.Kill());

            try
            {
                await XMLToDatabase();
            }
            catch(Exception ex)
            {
                Debug.Print(ex.StackTrace);
            }
            
            await Task.Run(() => OnProgressChanged.Report(false));
        }

        private static async Task XMLToDatabase()
        {
            //throw new NotImplementedException();
            FileInfo xFile = new FileInfo("Request_Definition.xml");
            if (!xFile.Exists) throw new Exception("XML File not exist.");

            await Task.Run(() =>
            {
                XElement xRequest = XElement.Load(xFile.FullName);
                //Debug.Print(xRequest.ToString());

                //string Method, Input Address $H$51, 
                //return value
                string xRange(string Address)
                {
                    IEnumerable<XElement> ieFilter =
                        from XElement x in xRequest.Elements()
                        where x.HasAttributes
                        select x;
                    IEnumerable<string> ieRangeVal =
                        from XElement x in ieFilter
                        where x.Attribute("CellAddress").Value == Address
                        && x.Attribute("Type").Value == "xlRange"
                        select x.Attribute("CellValue").Value;
                    int FoundNode = ieRangeVal.Count();
                    if (FoundNode != 1)
                        return "未入力";
                    return ieRangeVal.First();
                }
                //string Method, Input Checkbox Area $H$42:$K$43
                string xCheck(string Address)
                {
                    /*
                    Match mValid = Regex.Match(Address
                        , @"^\$[A-Z]\$\d+:\$[A-Z]\$\d+$");
                    if (!mValid.Success)
                        throw new Exception(Address + " is not a valid Cell Address Area.");
                    Regex rxExtract = new Regex(
                        @"^\$(?'CStart'[A-Z])\$(?'RStart'\d+):\$(?'CEnd'[A-Z])\$(?'REnd'\d+)$");
                    char CStart, CEnd;
                    int RStart, REnd;
                    CStart = rxExtract.Matches(Address)[0].Groups["CStart"].Value.ToCharArray()[0];
                    RStart = int.Parse(rxExtract.Matches(Address)[0].Groups["RStart"].Value);
                    CEnd = rxExtract.Matches(Address)[0].Groups["CEnd"].Value.ToCharArray()[0];
                    REnd = int.Parse(rxExtract.Matches(Address)[0].Groups["REnd"].Value);
                    //Debug.Print("CStart={0};CEnd={1};RStart={2};REnd={3}"
                    //    , CStart, CEnd, RStart, REnd);
                    string AddressRange = "";
                    for (char c = CStart; c <= CEnd; c++)
                    {
                        for (int i = RStart; i <= REnd; i++)
                        {
                            string thisRange = string.Format("${0}${1}", c, i);
                            AddressRange += thisRange + ";";
                        }
                    }
                    */
                    //$H$25;$H$26;$H$27;...$L$74;$L$75;$L$76;
                    CellAddresses TargetCellArea =
                    new CellAddresses(Address);
                    string AddressRange = TargetCellArea.FullAddress();
                    //Check if any of that is in the 
                    IEnumerable<XElement> ieFilter =
                        from XElement x in xRequest.Elements()
                        where x.HasAttributes
                        select x;
                    //Debug.Print(AddressRange);
                    IEnumerable<string> ieCheckVal =
                        from XElement x in ieFilter
                        where AddressRange.Contains(x.Attribute("CellAddress").Value)
                            && x.Attribute("Type").Value == "CheckBox"
                        select x.Attribute("CellValue").Value;
                    int ValueFound = ieCheckVal.Count();
                    switch (ValueFound)
                    {
                        case 0:
                            return "未入力";
                        case 1:
                            return ieCheckVal.First();
                        default:
                            return "無効な入力";
                    }
                }
                // Get Cell Value: xRange("$H$8"));
                // Get Checkbox Value: xCheck("$H$34:$K$36"));
                // Connect to the Database and sync it
                using (OleDbConnection conn = DatabaseConnection())
                {
                    conn.StateChange += (object sender, StateChangeEventArgs e)
                        => Debug.Print("Database Status Changed from {0} to {1}"
                        , e.OriginalState, e.CurrentState);
                    OleDbCommand INSERT_Host = new OleDbCommand();
                    OleDbTransaction TransactAll = null;
                    try
                    {
                        conn.Open();
                        TransactAll = conn.BeginTransaction(IsolationLevel.ReadCommitted);
                        INSERT_Host.Connection = conn;
                        INSERT_Host.Transaction = TransactAll;
                        //Command is ready.
                        
                        void MakeCommandforMA(string StartAddress)
                        {
                            Regex ValidHostname = new Regex(@"(\w|\d){8,}\.?");
                            CellAddress caStart = new CellAddress(StartAddress);
                            string C1VIP = xRange(caStart.FullAddress());
                            Match mhValidHost = ValidHostname.Match(C1VIP);
                            if (mhValidHost.Success)
                            {
                                INSERT_Host.CommandText = File.ReadAllText(SQLFolder + "INSERTHost.sql");
                                INSERT_Host.Parameters.AddWithValue("@ostname", C1VIP);
                                INSERT_Host.Parameters.AddWithValue("@PAddress", xRange(caStart.NextRow));
                                if (caStart.IsVIP)
                                {
                                    INSERT_Host.Parameters.AddWithValue("@aker", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@odel", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@PUCount", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@PUMicroprocessor", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@S", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@ersion", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@itVal", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@lusterBox", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@lusterIndex", "VIP");
                                }
                                else
                                {
                                    INSERT_Host.Parameters.AddWithValue("@aker", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@odel", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@PUCount", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@PUMicroprocessor", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@S", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@ersion", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@itVal", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@lusterBox", "VIP");
                                    INSERT_Host.Parameters.AddWithValue("@lusterIndex", "VIP");

                                }
                                INSERT_Host.ExecuteNonQuery();

                            }
                        }

                        MakeCommandforMA("$H$49");

                    }
                    catch (Exception)
                    {

                    }

                }


            });
        }

        private static async Task ExcelToXML(FileInfo ExcelFile, EXCEL.Worksheet xlSheet)
        {
            //throw new NotImplementedException();
            XElement xeRoot = new XElement("NewRequest");
            xeRoot.Add(new XElement("xlname", ExcelFile.Name));
            EXCEL.Range SyncRange = null;
            void SyncRangeToXML(string Address)
            {
                try
                {
                    SyncRange = xlSheet.Range[Address];
                    XElement xRange = new XElement(string.Format("R_{0}"
                        , SyncRange.Address.Replace("$", "")));
                    xRange.Add(new XAttribute("Type", "xlRange"));
                    xRange.Add(new XAttribute("CellAddress", SyncRange.Address));
                    xRange.Add(new XAttribute("CellValue", SyncRange.Value));

                    xeRoot.Add(xRange);
                }
                catch (Exception)
                {
                    Debug.Print(Address + " Cannot sync to XML.");
                }
            }
            // Sync CheckBoxes
            await Task.Run(() =>
            {
                IEnumerable<EXCEL.Shape> xlCheckboxes =
                    from EXCEL.Shape s in xlSheet.Shapes
                    where Regex.IsMatch(s.Name, @"(チェック|Check)")
                    && s.OLEFormat.Object.Value == 1
                    select s;
                foreach (EXCEL.Shape s in xlCheckboxes)
                {
                    XElement xCheckbox = new XElement(string.Format("C_{0}"
                        , s.TopLeftCell.Address.Replace("$", "")));
                    xCheckbox.Add(new XAttribute("Type", "CheckBox"));
                    xCheckbox.Add(new XAttribute("CellAddress", s.TopLeftCell.Address));
                    xCheckbox.Add(new XAttribute("CellValue", s.TopLeftCell.Offset[0, 1].Value));
                    xeRoot.Add(xCheckbox);
                }
            });

            for (int i = 7; i <= 11; i++)
                SyncRangeToXML("$H$" + i);
            for (int i = 32; i <= 100; i++)
                SyncRangeToXML("$H$" + i);
            for (int i = 32; i <= 100; i++)
                SyncRangeToXML("$L$" + i);
            for (int i = 32; i <= 100; i++)
                SyncRangeToXML("$P$" + i);

            SyncRangeToXML("$E$161");

            //Debug.Print(xeRoot.ToString());
            File.WriteAllText("Request_Definition.xml", xeRoot.ToString());
            Debug.Print ("File written Request_Definition.xml.");
            //Perfect, it gets the Values.
        }

        //
        // Summary:
        //     Check if a Request already existed in the database.
        private static bool CheckDatabaseExcelExist(string name)
        {
            //throw new NotImplementedException();
            var Conn = DatabaseConnection();
            OleDbCommand SELECT_CheckExcelExist = new OleDbCommand
            {
                CommandText =
                File.ReadAllText(SQLFolder + "CheckExist.sql"),
                Connection = Conn,
            };
            Conn.Open();
            SELECT_CheckExcelExist
                .Parameters.AddWithValue("@xlName", name);
            Debug.Print(SELECT_CheckExcelExist.CommandText);
            OleDbDataReader readerCheckExist =
                SELECT_CheckExcelExist.ExecuteReader();
            bool result = !readerCheckExist.HasRows; //If record has rows, it means not new
            Conn.Close(); Conn.Dispose();
            return result;
        }

        private static async Task MainExcelToDatabase(
            FileInfo FileName, EXCEL.Worksheet xlSheet)
        {
            //throw new NotImplementedException();
            await Task.Run(() => Debug.Print(xlSheet
                .Range["$H$51"].Value as string));
            using (OleDbConnection conn = DatabaseConnection())
            {
                conn.StateChange += (object sender, StateChangeEventArgs e)
                    => Debug.Print(string.Format
                    ("Database Status Changed from {0} to {1}."
                    , e.OriginalState, e.CurrentState));
                OleDbCommand INSERT_Host = new OleDbCommand();
                OleDbTransaction TransAll = null;
                try
                {
                    conn.Open();
                    TransAll = conn.BeginTransaction(IsolationLevel.ReadCommitted);
                    //Otherwise it throws the connection waiting for a local Transaction error
                    INSERT_Host.Connection = conn;
                    INSERT_Host.Transaction = TransAll;
                    Regex ValidHostname = 
                        new Regex(@"(\w|\d){8,}\.?");
                    //Is there Value to be synced at H49:K50, VIP, C1
                    string C1VIP = xlSheet.Range["$H$49"].Value;
                    Match mhValidHost = ValidHostname.Match(C1VIP);
                    if (mhValidHost.Success)
                    {
                        INSERT_Host.CommandText =File.ReadAllText(SQLFolder + "INSERTHost.sql");
                        INSERT_Host.Parameters.AddWithValue("@ostname", C1VIP);
                    }
                    string C2VIP = xlSheet.Range["$L$49"].Value;
                    mhValidHost = ValidHostname.Match(C2VIP);
                    if (mhValidHost.Success)
                    {
                        INSERT_Host.CommandText = File.ReadAllText(SQLFolder + "INSERTHost.sql");
                        INSERT_Host.Parameters.AddWithValue("@ostname", C2VIP);
                        INSERT_Host.Parameters.AddWithValue("@PAddress", xlSheet.Range["$L$50"].Value);
                        INSERT_Host.Parameters.AddWithValue("@aker", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@odel", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@PUCount", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@PUMicroprocessor", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@S", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@ersion", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@itVal", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@lusterBox", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@lusterIndex", "VIP");
                        INSERT_Host.ExecuteNonQuery();
                    }
                    string C3VIP = xlSheet.Range["$P$49"].Value;
                    mhValidHost = ValidHostname.Match(C3VIP);
                    if (mhValidHost.Success)
                    {
                        INSERT_Host.CommandText = File.ReadAllText(SQLFolder + "INSERTHost.sql");
                        INSERT_Host.Parameters.AddWithValue("@ostname", C3VIP);
                        INSERT_Host.Parameters.AddWithValue("@PAddress", xlSheet.Range["$P$50"].Value);
                        INSERT_Host.Parameters.AddWithValue("@aker", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@odel", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@PUCount", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@PUMicroprocessor", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@S", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@ersion", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@itVal", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@lusterBox", "VIP");
                        INSERT_Host.Parameters.AddWithValue("@lusterIndex", "VIP");
                        INSERT_Host.ExecuteNonQuery();
                    }
                    string C1PRI = xlSheet.Range["$H$51"].Value;
                    mhValidHost = ValidHostname.Match(C1PRI);
                    if (mhValidHost.Success)
                    {
                        INSERT_Host.CommandText = File.ReadAllText(SQLFolder + "INSERTHost.sql");
                        INSERT_Host.Parameters.AddWithValue("@ostname", C1PRI);
                        INSERT_Host.Parameters.AddWithValue("@PAddress", xlSheet.Range["$H$52"].Value);
                        INSERT_Host.Parameters.AddWithValue("@aker", xlSheet.Range["$H$53"].Value);
                        INSERT_Host.Parameters.AddWithValue("@odel", xlSheet.Range["$H$54"].Value);
                        INSERT_Host.Parameters.AddWithValue("@PUCount", xlSheet.Range["$H$55"].Value);
                        INSERT_Host.Parameters.AddWithValue("@PUMicroprocessor", xlSheet.Range["$H$56"].Value);
                        //Getting Check Box Info $H$57:$K$59
                        IEnumerable<EXCEL.Range> ieCheckBoxValue =
                            from EXCEL.Shape s in xlSheet.Shapes
                            where Regex.IsMatch(s.TopLeftCell.Address, @"^\$(H|K)\$5(7-9)$")
                            && s.OLEFormat.Object.Value == 1
                            && s.Name.Contains("Check Box")
                            select s.TopLeftCell.Offset[0, 1];

                        INSERT_Host.Parameters.AddWithValue("@S", ieCheckBoxValue.First().Value);
                        INSERT_Host.Parameters.AddWithValue("@ersion", xlSheet.Range["$H$60"].Value);
                        INSERT_Host.Parameters.AddWithValue("@itVal", xlSheet.Range["$H$61"].Value);
                        INSERT_Host.Parameters.AddWithValue("@lusterBox", xlSheet.Range["$H$62"].Value);
                        INSERT_Host.Parameters.AddWithValue("@lusterIndex", xlSheet.Range["$H$63"].Value);
                        INSERT_Host.ExecuteNonQuery();
                    }
                    TransAll.Commit();
                    Debug.Print("Sync Completed.");

                }
                catch(Exception ex)
                {

                    Debug.Print("An Error Happened, Rolling Back.");
                    
                    Debug.Print(ex.Message);
                    Debug.Print(ex.StackTrace);
                    TransAll.Rollback();
                }
            }

            /*
            string strSQL = File.ReadAllText(SQLFolder 
                + "INSERTtbExcel.sql");
            //Debug.Print(strSQL);

            OleDbCommand INSERTrequest = new OleDbCommand(strSQL, conn);
            try
            {
                INSERTrequest.ExecuteNonQuery();
            }
            catch(Exception ex)
            {
                Debug.Print(ex.Message);
            }*/
        }
    }

}
 