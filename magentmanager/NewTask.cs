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

namespace magentmanager
{
    class NewTask
    {
        const string SQLFolder =
            @"C:\Users\MoChen\source\repos\magentmanager\magentmanager\database\";
        const string DatabaseSource =
            SQLFolder + "magentr.accdb";
        const string DatabaseProvider =
            "Microsoft.ACE.OleDb.12.0";

        internal static Task<FileInfo[]> TaskNewFileOpen()
        {
            //throw new NotImplementedException();
            List<FileInfo> listFiles = new List<FileInfo>();
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Title = "同期したいな M/Agent 申請書を選択してください";
            openFile.Multiselect = true;
            openFile.CheckFileExists = true;
            openFile.DefaultExt = ".xlsx;.xls";
            openFile.Filter = "Excel Worksheet (.xls;.xlsx)|*.xls;*xlsx";
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

        internal static async Task TaskExcelToDatabase(FileInfo[] xlFiles)
        {
            //throw new NotImplementedException();
            
           
            EXCEL.Application xlApp = new EXCEL.Application();
            EXCEL.Workbooks xlWorkbooks = xlApp.Workbooks;
            
            foreach (FileInfo f in xlFiles)
            {
                EXCEL.Workbook xlWbk =
                xlWorkbooks.Open(f.FullName);
                EXCEL.Worksheet xlSheet = xlWbk.ActiveSheet;
                await MainExcelToDatabase(f, xlSheet);

                xlWbk.Close(false, Missing.Value, Missing.Value);
            }
            xlWorkbooks.Close();
            xlApp.Quit();
            GC.Collect();
        }
        
        private static async Task MainExcelToDatabase(
            FileInfo FileName, EXCEL.Worksheet xlSheet)
        {
            //throw new NotImplementedException();
            await Task.Run(() => Debug.Print(xlSheet
                .Range["$H$51"].Value as string));
            OleDbConnectionStringBuilder odcb = 
                new OleDbConnectionStringBuilder
            {
                Provider = DatabaseProvider,
                DataSource = DatabaseSource
            };
            string strconn = odcb.ToString();
            OleDbConnection conn = new OleDbConnection(strconn);
            conn.StateChange += (object sender, StateChangeEventArgs e)
                => Debug.Print(string.Format
                ("Database Status Changed from {0} to {1}."
                , e.OriginalState, e.CurrentState));
            conn.Open();
            string strSQL = File.ReadAllText(SQLFolder 
                + "INSERTmagentRequestForm.sql");
            //Debug.Print(strSQL);

            OleDbCommand INSERTrequest = new OleDbCommand(strSQL, conn);
            void AddParamCount()
            {
                Debug.Print("Current Parameters: " + INSERTrequest.Parameters.Count);
            }
            void AddText(string ParamName, string CellAddress)
            {
                Debug.Print("Getting Text Value from: " + CellAddress);
                if (xlSheet.Range[CellAddress].Value == null)
                {
                    INSERTrequest.Parameters
                        .AddWithValue(ParamName, "未入力");
                }
                else
                {
                    INSERTrequest.Parameters
                        .AddWithValue(ParamName, xlSheet.Range[CellAddress].Value);
                }
                AddParamCount();
            }
            void AddDate(string ParamName, string CellAddress)
            {
                Debug.Print("Getting Date Value from: " + CellAddress);

                var TargetDate = xlSheet.Range[CellAddress].Value;
                //Debug.Print("Target Date Type: " + TargetDate.);
                    INSERTrequest.Parameters
                        .AddWithValue(ParamName, TargetDate);
                AddParamCount();

            }
            void AddCheckbox(EXCEL.Range rng)
            {

            }
            Debug.Print("Parameter Count: " + INSERTrequest.Parameters.Count);
            INSERTrequest.Parameters.AddWithValue("xlname", FileName.Name);
            AddDate("$H$7");    AddText("$H$8");    AddText("$H$9");    AddText("$H$10");   AddText("$H$11");
            AddText("$H$11");   AddText("$H$11");   AddText("$H$11");

            AddText("$H$32"); AddText("$H$34"); AddText("$H$37"); AddText("$H$38"); AddText("$H$39");
            AddText("$H$40"); AddText("$H$41"); AddText("$H$42"); AddText("$H$44"); AddText("$H$48");
            AddText("$H$49"); AddText("$H$50"); AddText("$H$51"); AddText("$H$52"); AddText("$H$53");
            AddText("$H$54"); AddText("$H$55"); AddText("$H$56"); AddText("$H$57"); AddText("$H$60");
            AddText("$H$61"); AddText("$H$62"); AddText("$H$63"); AddText("$H$64"); AddText("$H$65");
            AddText("$H$66"); AddText("$H$67"); AddText("$H$68"); AddText("$H$69"); AddText("$H$70");
            AddText("$H$73"); AddText("$H$74"); AddText("$H$75"); AddText("$H$76"); AddText("$H$77");
            AddDate("$H$78"); AddDate("$H$79"); AddDate("$H$80"); AddText("$H$81"); AddText("$H$82");
            AddText("$H$83"); AddText("$H$84"); AddText("$H$85"); AddText("$H$86"); AddDate("$H$87");
            AddText("$H$88"); AddText("$H$89"); AddText("$H$90"); AddText("$H$91"); AddText("$H$92");
            AddText("$H$93"); AddText("$H$94"); AddText("$H$95"); AddText("$H$96"); AddText("$H$97");
            AddText("$H$98"); AddText("$H$99"); AddText("$H$100");

            AddText("$L$32"); AddText("$L$34"); AddText("$L$37"); AddText("$L$38"); AddText("$L$39");
            AddText("$L$40"); AddText("$L$41"); AddText("$L$42"); AddText("$L$44"); AddText("$L$48");
            AddText("$L$49"); AddText("$L$50"); AddText("$L$51"); AddText("$L$52"); AddText("$L$53");
            AddText("$L$54"); AddText("$L$55"); AddText("$L$56"); AddText("$L$57"); AddText("$L$60");
            AddText("$L$61"); AddText("$L$62"); AddText("$L$63"); AddText("$L$64"); AddText("$L$65");
            AddText("$L$66"); AddText("$L$67"); AddText("$L$68"); AddText("$L$69"); AddText("$L$70");
            AddText("$L$73"); AddText("$L$74"); AddText("$L$75"); AddText("$L$76"); AddText("$L$77");
            AddDate("$L$78"); AddDate("$L$79"); AddDate("$L$80"); AddText("$L$81"); AddText("$L$82");
            AddText("$L$83"); AddText("$L$84"); AddText("$L$85"); AddText("$L$86"); AddDate("$L$87");
            AddText("$L$88"); AddText("$L$89"); AddText("$L$90"); AddText("$L$91"); AddText("$L$92");
            AddText("$L$93"); AddText("$L$94"); AddText("$L$95"); AddText("$L$96"); AddText("$L$97");
            AddText("$L$98"); AddText("$L$99"); AddText("$L$100");

            AddText("$P$32"); AddText("$P$34"); AddText("$P$37"); AddText("$P$38"); AddText("$P$39");
            AddText("$P$40"); AddText("$P$41"); AddText("$P$42"); AddText("$P$44"); AddText("$P$48");
            AddText("$P$49"); AddText("$P$50"); AddText("$P$51"); AddText("$P$52"); AddText("$P$53");
            AddText("$P$54"); AddText("$P$55"); AddText("$P$56"); AddText("$P$57"); AddText("$P$60");
            AddText("$P$61"); AddText("$P$62"); AddText("$P$63"); AddText("$P$64"); AddText("$P$65");
            AddText("$P$66"); AddText("$P$67"); AddText("$P$68"); AddText("$P$69"); AddText("$P$70");
            AddText("$P$73"); AddText("$P$74"); AddText("$P$75"); AddText("$P$76"); AddText("$P$77");
            AddDate("$P$78"); AddDate("$P$79"); AddDate("$P$80"); AddText("$P$81"); AddText("$P$82");
            AddText("$P$83"); AddText("$P$84"); AddText("$P$85"); AddText("$P$86"); AddDate("$P$87");
            AddText("$P$88"); AddText("$P$89"); AddText("$P$90"); AddText("$P$91"); AddText("$P$92");
            AddText("$P$93"); AddText("$P$94"); AddText("$P$95"); AddText("$P$96"); AddText("$P$97");
            AddText("$P$98"); AddText("$P$99"); AddText("$P$100");

            AddText("$E$161");

            AddText("date_apply", "$H$7");
            AddText("applier", "$H$8");
            AddText("mailaddress", "$H$9");
            AddText("phonenumber", "$H$10");
            AddText("approver", "$H$11");
            AddText("jobcon_accept", "$H$11"); //Temp
            AddText("jobcon_confirm", "$H$11");//Temp
            AddText("jobcon_approve", "$H$11");//Temp
            AddText("classification1", "$H$32");
            AddText("change_point1", "$H$34");
            AddText("SIer1", "$H$37");
            AddText("serverPIC1", "$H$38");
            AddText("sysid1", "$H$39");
            AddText("sysname1", "$H$40");
            AddText("sysname_sub1", "$H$41");
            AddText("network_location1", "$H$42");
            AddText("network_area1", "$H$44");
            AddText("iscluster1", "$H$48");
            AddText("viphost1", "$H$49");
            AddText("vipaddress1", "$H$50");
            AddText("prihost1", "$H$51");
            AddText("priadress1", "$H$52");
            AddText("primaker1", "$H$53");
            AddText("primodel1", "$H$54");
            AddText("pricpunumber1", "$H$55");
            AddText("pricpumicro1", "$H$56");
            AddText("prios1", "$H$57");
            AddText("priversion1", "$H$60");
            AddText("pribit1", "$H$61");
            AddText("pribox1", "$H$62");
            AddText("priboxindex1", "$H$63");
            AddText("sechost1", "$H$64");
            AddText("secadress1", "$H$65");
            AddText("secmaker1", "$H$66");
            AddText("secmodel1", "$H$67");
            AddText("seccpunumber1", "$H$68");
            AddText("seccpumicro1", "$H$69");
            AddText("secos1", "$H$70");
            AddText("secversion1", "$H$73");
            AddText("secbit1", "$H$74");
            AddText("secbox1", "$H$75");
            AddText("secboxindex1", CellRange);
            AddText("mstmaport1", CellRange);
            AddText("date_install1", CellRange);
            AddText("date_connect1", CellRange);
            AddText("date_jobstart1", CellRange);
            AddText("jobnumer1", CellRange);
            AddText("hascallorder1", CellRange);
            AddText("hasfirewall1", CellRange);
            AddText("maversion1", CellRange);
            AddText("isfirst1", CellRange);
            AddText("isprod1", CellRange);
            AddText("date_qadone1", CellRange);
            AddText("costfrom1", CellRange);
            AddText("costsysname1", CellRange);
            AddText("costsysname_sub1", CellRange);
            AddText("hassunday1", CellRange);
            AddText("isrelated1", CellRange);
            AddText("relate_sysname1", CellRange);
            AddText("relate_sysname_sub1", CellRange);
            AddText("related_sysid1", CellRange);
            AddText("relate_ms1", CellRange);
            AddText("matmsport1", CellRange);
            AddText("vipms1", CellRange);
            AddText("prims1", CellRange);
            AddText("secms1", CellRange);
            AddText("classification2", CellRange);
            AddText("change_point2", CellRange);
            AddText("SIer2", CellRange);
            AddText("serverPIC2", CellRange);
            AddText("sysid2", CellRange);
            AddText("sysname2", CellRange);
            AddText("sysname_sub2", CellRange);
            AddText("network_location2", CellRange);
            AddText("network_area2", CellRange);
            AddText("iscluster2", CellRange);
            AddText("viphost2", CellRange);
            AddText("vipaddress2", CellRange);
            AddText("prihost2", CellRange);
            AddText("priadress2", CellRange);
            AddText("primaker2", CellRange);
            AddText("primodel2", CellRange);
            AddText("pricpunumber2", CellRange);
            AddText("pricpumicro2", CellRange);
            AddText("prios2", CellRange);
            AddText("priversion2", CellRange);
            AddText("pribit2", CellRange);
            AddText("pribox2", CellRange);
            AddText("priboxindex2", CellRange);
            AddText("sechost2", CellRange);
            AddText("secadress2", CellRange);
            AddText("secmaker2", CellRange);
            AddText("secmodel2", CellRange);
            AddText("seccpunumber2", CellRange);
            AddText("seccpumicro2", CellRange);
            AddText("secos2", CellRange);
            AddText("secversion2", CellRange);
            AddText("secbit2", CellRange);
            AddText("secbox2", CellRange);
            AddText("secboxindex2", CellRange);
            AddText("mstmaport2", CellRange);
            AddText("date_install2", CellRange);
            AddText("date_connect2", CellRange);
            AddText("date_jobstart2", CellRange);
            AddText("jobnumer2", CellRange);
            AddText("hascallorder2", CellRange);
            AddText("hasfirewall2", CellRange);
            AddText("maversion2", CellRange);
            AddText("isfirst2", CellRange);
            AddText("isprod2", CellRange);
            AddText("date_qadone2", CellRange);
            AddText("costfrom2", CellRange);
            AddText("costsysname2", CellRange);
            AddText("costsysname_sub2", CellRange);
            AddText("hassunday2", CellRange);
            AddText("isrelated2", CellRange);
            AddText("relate_sysname2", CellRange);
            AddText("relate_sysname_sub2", CellRange);
            AddText("related_sysid2", CellRange);
            AddText("relate_ms2", CellRange);
            AddText("matmsport2", CellRange);
            AddText("vipms2", CellRange);
            AddText("prims2", CellRange);
            AddText("secms2", CellRange);
            AddText("classification3", CellRange);
            AddText("change_point3", CellRange);
            AddText("SIer3", CellRange);
            AddText("serverPIC3", CellRange);
            AddText("sysid3", CellRange);
            AddText("sysname3", CellRange);
            AddText("sysname_sub3", CellRange);
            AddText("network_location3", CellRange);
            AddText("network_area3", CellRange);
            AddText("iscluster3", CellRange);
            AddText("viphost3", CellRange);
            AddText("vipaddress3", CellRange);
            AddText("prihost3", CellRange);
            AddText("priadress3", CellRange);
            AddText("primaker3", CellRange);
            AddText("primodel3", CellRange);
            AddText("pricpunumber3", CellRange);
            AddText("pricpumicro3", CellRange);
            AddText("prios3", CellRange);
            AddText("priversion3", CellRange);
            AddText("pribit3", CellRange);
            AddText("pribox3", CellRange);
            AddText("priboxindex3", CellRange);
            AddText("sechost3", CellRange);
            AddText("secadress3", CellRange);
            AddText("secmaker3", CellRange);
            AddText("secmodel3", CellRange);
            AddText("seccpunumber3", CellRange);
            AddText("seccpumicro3", CellRange);
            AddText("secos3", CellRange);
            AddText("secversion3", CellRange);
            AddText("secbit3", CellRange);
            AddText("secbox3", CellRange);
            AddText("secboxindex3", CellRange);
            AddText("mstmaport3", CellRange);
            AddText("date_install3", CellRange);
            AddText("date_connect3", CellRange);
            AddText("date_jobstart3", CellRange);
            AddText("jobnumer3", CellRange);
            AddText("hascallorder3", CellRange);
            AddText("hasfirewall3", CellRange);
            AddText("maversion3", CellRange);
            AddText("isfirst3", CellRange);
            AddText("isprod3", CellRange);
            AddText("date_qadone3", CellRange);
            AddText("costfrom3", CellRange);
            AddText("costsysname3", CellRange);
            AddText("costsysname_sub3", CellRange);
            AddText("hassunday3", CellRange);
            AddText("isrelated3", CellRange);
            AddText("relate_sysname3", CellRange);
            AddText("relate_sysname_sub3", CellRange);
            AddText("related_sysid3", CellRange);
            AddText("relate_ms3", CellRange);
            AddText("matmsport3", CellRange);
            AddText("vipms3", CellRange);
            AddText("prims3", CellRange);
            AddText("secms3", CellRange);
            AddText("specialcomment", CellRange);

            
                        Debug.Print("Parameter Count: " + INSERTrequest.Parameters.Count);
            Debug.Print(INSERTrequest.CommandText);
            /*
            INSERTrequest.Parameters.AddWithValue("jobcon_accept", Values);
            INSERTrequest.Parameters.AddWithValue("jobcon_confirm", Values);
            INSERTrequest.Parameters.AddWithValue("jobcon_approve", Values);
            INSERTrequest.Parameters.AddWithValue("classification1", Values);
            INSERTrequest.Parameters.AddWithValue("change_point1", Values);
            INSERTrequest.Parameters.AddWithValue("SIer1", Values);
            INSERTrequest.Parameters.AddWithValue("serverPIC1", Values);
            INSERTrequest.Parameters.AddWithValue("sysid1", Values);
            INSERTrequest.Parameters.AddWithValue("sysname1", Values);
            INSERTrequest.Parameters.AddWithValue("sysname_sub1", Values);
            INSERTrequest.Parameters.AddWithValue("network_location1", Values);
            INSERTrequest.Parameters.AddWithValue("network_area1", Values);
            INSERTrequest.Parameters.AddWithValue("iscluster1", Values);
            INSERTrequest.Parameters.AddWithValue("viphost1", Values);
            INSERTrequest.Parameters.AddWithValue("vipaddress1", Values);
            INSERTrequest.Parameters.AddWithValue("prihost1", Values);
            INSERTrequest.Parameters.AddWithValue("priadress1", Values);
            INSERTrequest.Parameters.AddWithValue("primaker1", Values);
            INSERTrequest.Parameters.AddWithValue("primodel1", Values);
            INSERTrequest.Parameters.AddWithValue("pricpunumber1", Values);
            INSERTrequest.Parameters.AddWithValue("pricpumicro1", Values);
            INSERTrequest.Parameters.AddWithValue("prios1", Values);
            INSERTrequest.Parameters.AddWithValue("priversion1", Values);
            INSERTrequest.Parameters.AddWithValue("pribit1", Values);
            INSERTrequest.Parameters.AddWithValue("pribox1", Values);
            INSERTrequest.Parameters.AddWithValue("priboxindex1", Values);
            INSERTrequest.Parameters.AddWithValue("sechost1", Values);
            INSERTrequest.Parameters.AddWithValue("secadress1", Values);
            INSERTrequest.Parameters.AddWithValue("secmaker1", Values);
            INSERTrequest.Parameters.AddWithValue("secmodel1", Values);
            INSERTrequest.Parameters.AddWithValue("seccpunumber1", Values);
            INSERTrequest.Parameters.AddWithValue("seccpumicro1", Values);
            INSERTrequest.Parameters.AddWithValue("secos1", Values);
            INSERTrequest.Parameters.AddWithValue("secversion1", Values);
            INSERTrequest.Parameters.AddWithValue("secbit1", Values);
            INSERTrequest.Parameters.AddWithValue("secbox1", Values);
            INSERTrequest.Parameters.AddWithValue("secboxindex1", Values);
            INSERTrequest.Parameters.AddWithValue("mstmaport1", Values);
            INSERTrequest.Parameters.AddWithValue("date_install1", Values);
            INSERTrequest.Parameters.AddWithValue("date_connect1", Values);
            INSERTrequest.Parameters.AddWithValue("date_jobstart1", Values);
            INSERTrequest.Parameters.AddWithValue("jobnumer1", Values);
            INSERTrequest.Parameters.AddWithValue("hascallorder1", Values);
            INSERTrequest.Parameters.AddWithValue("hasfirewall1", Values);
            INSERTrequest.Parameters.AddWithValue("maversion1", Values);
            INSERTrequest.Parameters.AddWithValue("isfirst1", Values);
            INSERTrequest.Parameters.AddWithValue("isprod1", Values);
            INSERTrequest.Parameters.AddWithValue("date_qadone1", Values);
            INSERTrequest.Parameters.AddWithValue("costfrom1", Values);
            INSERTrequest.Parameters.AddWithValue("costsysname1", Values);
            INSERTrequest.Parameters.AddWithValue("costsysname_sub1", Values);
            INSERTrequest.Parameters.AddWithValue("hassunday1", Values);
            INSERTrequest.Parameters.AddWithValue("isrelated1", Values);
            INSERTrequest.Parameters.AddWithValue("relate_sysname1", Values);
            INSERTrequest.Parameters.AddWithValue("relate_sysname_sub1", Values);
            INSERTrequest.Parameters.AddWithValue("relate_ms1", Values);
            INSERTrequest.Parameters.AddWithValue("matmsport1", Values);
            INSERTrequest.Parameters.AddWithValue("vipms1", Values);
            INSERTrequest.Parameters.AddWithValue("prims1", Values);
            INSERTrequest.Parameters.AddWithValue("secms1", Values);
            INSERTrequest.Parameters.AddWithValue("classification2", Values);
            INSERTrequest.Parameters.AddWithValue("change_point2", Values);
            INSERTrequest.Parameters.AddWithValue("SIer2", Values);
            INSERTrequest.Parameters.AddWithValue("serverPIC2", Values);
            INSERTrequest.Parameters.AddWithValue("sysid2", Values);
            INSERTrequest.Parameters.AddWithValue("sysname2", Values);
            INSERTrequest.Parameters.AddWithValue("sysname_sub2", Values);
            INSERTrequest.Parameters.AddWithValue("network_location2", Values);
            INSERTrequest.Parameters.AddWithValue("network_area2", Values);
            INSERTrequest.Parameters.AddWithValue("iscluster2", Values);
            INSERTrequest.Parameters.AddWithValue("viphost2", Values);
            INSERTrequest.Parameters.AddWithValue("vipaddress2", Values);
            INSERTrequest.Parameters.AddWithValue("prihost2", Values);
            INSERTrequest.Parameters.AddWithValue("priadress2", Values);
            INSERTrequest.Parameters.AddWithValue("primaker2", Values);
            INSERTrequest.Parameters.AddWithValue("primodel2", Values);
            INSERTrequest.Parameters.AddWithValue("pricpunumber2", Values);
            INSERTrequest.Parameters.AddWithValue("pricpumicro2", Values);
            INSERTrequest.Parameters.AddWithValue("prios2", Values);
            INSERTrequest.Parameters.AddWithValue("priversion2", Values);
            INSERTrequest.Parameters.AddWithValue("pribit2", Values);
            INSERTrequest.Parameters.AddWithValue("pribox2", Values);
            INSERTrequest.Parameters.AddWithValue("priboxindex2", Values);
            INSERTrequest.Parameters.AddWithValue("sechost2", Values);
            INSERTrequest.Parameters.AddWithValue("secadress2", Values);
            INSERTrequest.Parameters.AddWithValue("secmaker2", Values);
            INSERTrequest.Parameters.AddWithValue("secmodel2", Values);
            INSERTrequest.Parameters.AddWithValue("seccpunumber2", Values);
            INSERTrequest.Parameters.AddWithValue("seccpumicro2", Values);
            INSERTrequest.Parameters.AddWithValue("secos2", Values);
            INSERTrequest.Parameters.AddWithValue("secversion2", Values);
            INSERTrequest.Parameters.AddWithValue("secbit2", Values);
            INSERTrequest.Parameters.AddWithValue("secbox2", Values);
            INSERTrequest.Parameters.AddWithValue("secboxindex2", Values);
            INSERTrequest.Parameters.AddWithValue("mstmaport2", Values);
            INSERTrequest.Parameters.AddWithValue("date_install2", Values);
            INSERTrequest.Parameters.AddWithValue("date_connect2", Values);
            INSERTrequest.Parameters.AddWithValue("date_jobstart2", Values);
            INSERTrequest.Parameters.AddWithValue("jobnumer2", Values);
            INSERTrequest.Parameters.AddWithValue("hascallorder2", Values);
            INSERTrequest.Parameters.AddWithValue("hasfirewall2", Values);
            INSERTrequest.Parameters.AddWithValue("maversion2", Values);
            INSERTrequest.Parameters.AddWithValue("isfirst2", Values);
            INSERTrequest.Parameters.AddWithValue("isprod2", Values);
            INSERTrequest.Parameters.AddWithValue("date_qadone2", Values);
            INSERTrequest.Parameters.AddWithValue("costfrom2", Values);
            INSERTrequest.Parameters.AddWithValue("costsysname2", Values);
            INSERTrequest.Parameters.AddWithValue("costsysname_sub2", Values);
            INSERTrequest.Parameters.AddWithValue("hassunday2", Values);
            INSERTrequest.Parameters.AddWithValue("isrelated2", Values);
            INSERTrequest.Parameters.AddWithValue("relate_sysname2", Values);
            INSERTrequest.Parameters.AddWithValue("relate_sysname_sub2", Values);
            INSERTrequest.Parameters.AddWithValue("relate_ms2", Values);
            INSERTrequest.Parameters.AddWithValue("matmsport2", Values);
            INSERTrequest.Parameters.AddWithValue("vipms2", Values);
            INSERTrequest.Parameters.AddWithValue("prims2", Values);
            INSERTrequest.Parameters.AddWithValue("secms2", Values);
            INSERTrequest.Parameters.AddWithValue("classification3", Values);
            INSERTrequest.Parameters.AddWithValue("change_point3", Values);
            INSERTrequest.Parameters.AddWithValue("SIer3", Values);
            INSERTrequest.Parameters.AddWithValue("serverPIC3", Values);
            INSERTrequest.Parameters.AddWithValue("sysid3", Values);
            INSERTrequest.Parameters.AddWithValue("sysname3", Values);
            INSERTrequest.Parameters.AddWithValue("sysname_sub3", Values);
            INSERTrequest.Parameters.AddWithValue("network_location3", Values);
            INSERTrequest.Parameters.AddWithValue("network_area3", Values);
            INSERTrequest.Parameters.AddWithValue("iscluster3", Values);
            INSERTrequest.Parameters.AddWithValue("viphost3", Values);
            INSERTrequest.Parameters.AddWithValue("vipaddress3", Values);
            INSERTrequest.Parameters.AddWithValue("prihost3", Values);
            INSERTrequest.Parameters.AddWithValue("priadress3", Values);
            INSERTrequest.Parameters.AddWithValue("primaker3", Values);
            INSERTrequest.Parameters.AddWithValue("primodel3", Values);
            INSERTrequest.Parameters.AddWithValue("pricpunumber3", Values);
            INSERTrequest.Parameters.AddWithValue("pricpumicro3", Values);
            INSERTrequest.Parameters.AddWithValue("prios3", Values);
            INSERTrequest.Parameters.AddWithValue("priversion3", Values);
            INSERTrequest.Parameters.AddWithValue("pribit3", Values);
            INSERTrequest.Parameters.AddWithValue("pribox3", Values);
            INSERTrequest.Parameters.AddWithValue("priboxindex3", Values);
            INSERTrequest.Parameters.AddWithValue("sechost3", Values);
            INSERTrequest.Parameters.AddWithValue("secadress3", Values);
            INSERTrequest.Parameters.AddWithValue("secmaker3", Values);
            INSERTrequest.Parameters.AddWithValue("secmodel3", Values);
            INSERTrequest.Parameters.AddWithValue("seccpunumber3", Values);
            INSERTrequest.Parameters.AddWithValue("seccpumicro3", Values);
            INSERTrequest.Parameters.AddWithValue("secos3", Values);
            INSERTrequest.Parameters.AddWithValue("secversion3", Values);
            INSERTrequest.Parameters.AddWithValue("secbit3", Values);
            INSERTrequest.Parameters.AddWithValue("secbox3", Values);
            INSERTrequest.Parameters.AddWithValue("secboxindex3", Values);
            INSERTrequest.Parameters.AddWithValue("mstmaport3", Values);
            INSERTrequest.Parameters.AddWithValue("date_install3", Values);
            INSERTrequest.Parameters.AddWithValue("date_connect3", Values);
            INSERTrequest.Parameters.AddWithValue("date_jobstart3", Values);
            INSERTrequest.Parameters.AddWithValue("jobnumer3", Values);
            INSERTrequest.Parameters.AddWithValue("hascallorder3", Values);
            INSERTrequest.Parameters.AddWithValue("hasfirewall3", Values);
            INSERTrequest.Parameters.AddWithValue("maversion3", Values);
            INSERTrequest.Parameters.AddWithValue("isfirst3", Values);
            INSERTrequest.Parameters.AddWithValue("isprod3", Values);
            INSERTrequest.Parameters.AddWithValue("date_qadone3", Values);
            INSERTrequest.Parameters.AddWithValue("costfrom3", Values);
            INSERTrequest.Parameters.AddWithValue("costsysname3", Values);
            INSERTrequest.Parameters.AddWithValue("costsysname_sub3", Values);
            INSERTrequest.Parameters.AddWithValue("hassunday3", Values);
            INSERTrequest.Parameters.AddWithValue("isrelated3", Values);
            INSERTrequest.Parameters.AddWithValue("relate_sysname3", Values);
            INSERTrequest.Parameters.AddWithValue("relate_sysname_sub3", Values);
            INSERTrequest.Parameters.AddWithValue("relate_ms3", Values);
            INSERTrequest.Parameters.AddWithValue("matmsport3", Values);
            INSERTrequest.Parameters.AddWithValue("vipms3", Values);
            INSERTrequest.Parameters.AddWithValue("prims3", Values);
            INSERTrequest.Parameters.AddWithValue("secms3", Values);
            INSERTrequest.Parameters.AddWithValue("specialcomment, Values);*/
            try
            {
                INSERTrequest.ExecuteNonQuery();
            }
            catch(Exception ex)
            {
                Debug.Print(ex.Message);
            }
        }
    }
}
