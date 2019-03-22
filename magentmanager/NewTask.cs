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

        //Sync Excel File to a Database.
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
                        await MainExcelToDatabase(f, xlSheet);
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
            Console.Beep(750, 50);
            await Task.Run(() => CurrentExcelProcess.Kill());
            await Task.Run(() => OnProgressChanged.Report(false));
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
            OleDbConnection conn = DatabaseConnection();
            conn.StateChange += (object sender, StateChangeEventArgs e)
                => Debug.Print(string.Format
                ("Database Status Changed from {0} to {1}."
                , e.OriginalState, e.CurrentState));
            conn.Open();

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
            conn.Close();
        }
    }
}
 