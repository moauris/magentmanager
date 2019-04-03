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
                DateTime xDate(string Address)
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
                        return DateTime.Parse("1900-01-01");
                    string DateString = ieRangeVal.First();
                    bool IsDate = DateTime.TryParse(DateString, out DateTime result);
                    if (IsDate)
                    {
                        return result;
                    }
                    else
                    {
                        return DateTime.Parse("1900-01-01");
                    }
                }
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
                //No, there is no need 
                string xCheck(string Address)
                {
                    CellRange TargetCellArea =
                    new CellRange(Address);
                    string AddressRange = TargetCellArea.Extend();
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
                    OleDbCommand INSERT_Excel = new OleDbCommand();
                    OleDbCommand INSERT_Agent = new OleDbCommand();

                    OleDbTransaction TransactAll = null;
                    conn.Open();
                    TransactAll = conn.BeginTransaction(IsolationLevel.ReadCommitted);
                    INSERT_Host.Connection = conn;
                    INSERT_Host.Transaction = TransactAll;
                    INSERT_Excel = INSERT_Agent = INSERT_Host.Clone();

                    //Make 3 Commands, one for each column
                    //return value is rows affected.
                    int MakeCommandForServer(string StartAddress)
                    {
                        Regex ValidHostname = new Regex(@"(\w|\d){8,}\.?");
                        CellRange caStart = new CellRange(StartAddress);
                        string MA_Hostname = xRange(caStart.Extend());
                        Match mhValidHost = ValidHostname.Match(MA_Hostname);
                        if (mhValidHost.Success)
                        {
                            INSERT_Host.CommandText = File.ReadAllText(SQLFolder + "INSERTHost.sql");
                            INSERT_Host.Parameters.AddWithValue("@ostname", MA_Hostname);
                            INSERT_Host.Parameters.AddWithValue("@PAddress", xRange(caStart.NextRow()));
                            if (caStart.InitRange().Contains("49")) //M/Agent Range at 49 is VIP
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
                                INSERT_Host.Parameters.AddWithValue("@aker", xRange(caStart.NextRow()));
                                INSERT_Host.Parameters.AddWithValue("@odel", xRange(caStart.NextRow()));
                                INSERT_Host.Parameters.AddWithValue("@PUCount", xRange(caStart.NextRow()));
                                INSERT_Host.Parameters.AddWithValue("@PUMicroprocessor", xRange(caStart.NextRow()));
                                INSERT_Host.Parameters.AddWithValue("@S", xCheck(caStart.NextArea(3)));
                                INSERT_Host.Parameters.AddWithValue("@ersion", xRange(caStart.NextRow()));
                                INSERT_Host.Parameters.AddWithValue("@itVal", xRange(caStart.NextRow()));
                                INSERT_Host.Parameters.AddWithValue("@lusterBox", xRange(caStart.NextRow()));
                                INSERT_Host.Parameters.AddWithValue("@lusterIndex", xRange(caStart.NextRow())); 

                            }
                            //Debug.Print("Parameters Before Clear: {0}", INSERT_Host.Parameters.Count);
                            int RowsAffected = INSERT_Host.ExecuteNonQuery();
                            INSERT_Host.Parameters.Clear();
                            return RowsAffected;
                        }
                        return 0;
                    }
                    int MakeCommandforAgent(char column)
                    {
                        if (!char.IsLetter(column))
                            throw new Exception("Invalid Column Character");
                        Regex ValidHostname = new Regex(@"(\w|\d){8,}\.?");

                        INSERT_Agent.CommandText = File.ReadAllText(SQLFolder + "INSERTAgent.sql");
                        //Agent Name is like uny40310.abc00101
                        //If VIP Exist, use VIP; If VIP not exist, use Pri
                        //Validate VIP At 49, and PRI at 51
                        string magentName = "";
                        string tempRange = string.Format(@"${0}$49", column);
                        if (ValidHostname.IsMatch(xRange(tempRange)))
                            magentName = xRange(tempRange);
                        tempRange = string.Format(@"${0}$51", column);
                        if (ValidHostname.IsMatch(xRange(tempRange)))
                            magentName = xRange(tempRange);
                        if(magentName == "")
                            throw new Exception("Invalid MA Column, magentName Failed to caputre.");

                        tempRange = string.Format(@"${0}$98", column);
                        string mserverName = xRange(tempRange);
                        string AgentName = string.Format(@"{0}.{1}", mserverName, magentName);

                        INSERT_Agent.Parameters.AddWithValue("@AgentName", AgentName);
                        INSERT_Agent.Parameters.AddWithValue("@rlnFileName", xRequest.Element("xlname").Value);

                        tempRange = string.Format(@"${0}$31", column);
                        CellRange cell = new CellRange(tempRange);
                        INSERT_Agent.Parameters.AddWithValue("@ApplyType", xCheck(cell.NextArea(2)));
                        INSERT_Agent.Parameters.AddWithValue("@ChangePoint", xCheck(cell.NextArea(3)));
                        INSERT_Agent.Parameters.AddWithValue("@SIer", xRange(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@ServerPIC", xRange(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@SystemID", xRange(cell.NextRow())); //7

                        INSERT_Agent.Parameters.AddWithValue("@SystemName", xRange(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@SystemSubName", xRange(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@NetworkLocation", xCheck(cell.NextArea(2)));
                        INSERT_Agent.Parameters.AddWithValue("@NetworkArea", xCheck(cell.NextArea(4))); //11

                        INSERT_Agent.Parameters.AddWithValue("@ServerVIP", xRange(cell.NextRow(2)));
                        INSERT_Agent.Parameters.AddWithValue("@ServerPRI", xRange(cell.NextRow(2)));
                        INSERT_Agent.Parameters.AddWithValue("@ServerSEC", xRange(cell.NextRow(13))); //14

                        INSERT_Agent.Parameters.AddWithValue("@MStMACommunicationPort", xRange(cell.NextRow(13)));
                        INSERT_Agent.Parameters.AddWithValue("@MA_InstallDate", xDate(cell.NextRow())); //16

                        INSERT_Agent.Parameters.AddWithValue("@MS_Connection", xDate(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@JobStartDate", xDate(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@JobCount", xRange(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@HasCallorder", xCheck(cell.NextArea(1)));

                        INSERT_Agent.Parameters.AddWithValue("@HasFirewall", xCheck(cell.NextArea(1))); //21

                        INSERT_Agent.Parameters.AddWithValue("@MA_Version", xRange(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@IsFirstTime", xCheck(cell.NextArea(1)));
                        INSERT_Agent.Parameters.AddWithValue("@IsProduction", xCheck(cell.NextArea(1)));
                        INSERT_Agent.Parameters.AddWithValue("@TestDoneDate", xDate(cell.NextRow())); //25

                        INSERT_Agent.Parameters.AddWithValue("@CostFrom", xRange(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@CostFromSystemName", xRange(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@CostFromSubSystemName", xRange(cell.NextRow())); //28

                        INSERT_Agent.Parameters.AddWithValue("@HasSundayJobs", xCheck(cell.NextArea(1)));
                        INSERT_Agent.Parameters.AddWithValue("@HasRelatedSystems", xCheck(cell.NextArea(1)));
                        INSERT_Agent.Parameters.AddWithValue("@RelatedSystemID", xRange(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@RelatedSystemName", xRange(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@RelatedSystemSubName", xRange(cell.NextRow()));


                        INSERT_Agent.Parameters.AddWithValue("@RelatedSystemDatacenter", xRange(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@MAtMSCommunicationPort", xRange(cell.NextRow())); //30

                        INSERT_Agent.Parameters.AddWithValue("@MSVIP", xRange(cell.NextRow())); 
                        INSERT_Agent.Parameters.AddWithValue("@MSPRI", xRange(cell.NextRow()));
                        INSERT_Agent.Parameters.AddWithValue("@MSSEC", xRange(cell.NextRow())); //33
                        int RowsAffected = INSERT_Agent.ExecuteNonQuery();
                        INSERT_Agent.Parameters.Clear();
                        return RowsAffected;
                        /*
                         @AgentName, @rlnFileName, @ApplyType, @ChangePoint, @SIer, @ServerPIC,
                         @SystemID, @SystemName, @SystemSubName, @NetworkLocation, @NetworkArea,
                         @ServerVIP, @ServerPRI, @ServerSEC, @MStMACommunicationPort, @MA_InstallDate,
                         @MS_Connection, @JobStartDate, @JobCount, @HasCallorder, @HasFirewall, @MA_Version,
                         @IsFirstTime, @IsProduction, @TestDoneDate, @CostFrom, @CostFromSystemName,
                         @CostFromSubSystemName, @HasSundayJobs, @HasRelatedSystems, @RelatedSystemID,
                         @RelatedSystemName, @RelatedSystemSubName, @RelatedSystemDatacenter,
                         @MAtMSCommunicationPort, @MSVIP, @MSPRI, @MSSEC 38
                         */
                    }
                    // Sync the information of the new task
                    // return values is rows affected
                    int MakeCommandforExcel()
                    {
                        INSERT_Excel.CommandText = File.ReadAllText(SQLFolder + "INSERTExcel.sql");
                        INSERT_Excel.Parameters.AddWithValue("@lname", xRequest.Element("xlname").Value);
                        CellRange caStart = new CellRange("$H$7");
                        string applydate = xRange(caStart.EndinRange());
                        Debug.Print(applydate);
                        INSERT_Excel.Parameters.AddWithValue("@ate_apply", DateTime.Parse(applydate));
                        INSERT_Excel.Parameters.AddWithValue("@pplier", xRange(caStart.NextRow()));
                        INSERT_Excel.Parameters.AddWithValue("@ailaddress", xRange(caStart.NextRow()));
                        INSERT_Excel.Parameters.AddWithValue("@honenumber", xRange(caStart.NextRow()));
                        INSERT_Excel.Parameters.AddWithValue("@pprover", xRange(caStart.NextRow()));
                        INSERT_Excel.Parameters.AddWithValue("@obcon_accept", "陳 黙");
                        INSERT_Excel.Parameters.AddWithValue("@obcon_confirm", "徐 長練");
                        INSERT_Excel.Parameters.AddWithValue("@obcon_approve", "孫 紅莉");
                        INSERT_Excel.Parameters.AddWithValue("@pecialcomment", xRange("$E$161"));
                        Debug.Print("INSERT_Excel Parameters Count: {0}", INSERT_Excel.Parameters.Count);
                        //Debug.Print(INSERT_Excel.CommandText);
                        return INSERT_Excel.ExecuteNonQuery();
                    }//(@lname, @ate_apply, @pplier, @ailaddress, @honenumber, @pprover, @obcon_accept, @obcon_confirm, @obcon_approve, @pecialcomment);

                    if (MakeCommandForServer("$H$49") == 1 || MakeCommandForServer("$H$51") == 1 || MakeCommandForServer("$H$64") == 1)
                    {
                        MakeCommandforAgent('H');
                    }

                    if (MakeCommandForServer("$L$49") == 1 || MakeCommandForServer("$L$51") == 1 || MakeCommandForServer("$L$64") == 1)
                    {
                        MakeCommandforAgent('L');
                    }
                    if (MakeCommandForServer("$P$49") == 1 || MakeCommandForServer("$P$51") == 1 || MakeCommandForServer("$P$64") == 1)
                    {
                        MakeCommandforAgent('P');
                    }

                    MakeCommandforExcel();

                    TransactAll.Commit();
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
 