using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
//using MySql.Data.MySqlClient;
using System.Data;
using System.Data.SqlClient;

namespace RestoreMPN
{
    class Program
    {
        static int TIDIn=0, TIDOut=0;
        static string commandIn ="",commandOut="", commandmssqlIn,commandmssqlOut;
        static string fileName = "",folderName = "",dataBaseSource="";
        static System.IO.StreamReader file;
        static int failed = 0;
        //static string sqlstring = "SERVER=127.0.0.1;PORT=3306;DATABASE=ffdb;UID=root;PASSWORD=;";
        static string mssqlstring = "Data Source=(localdb)\\ProjectsV13;Initial Catalog=FFDB;Integrated Security=True;Connect Timeout=60;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
        //static string mssqlstring = "Server=LAMBCHOP;Database=FFSQLDB_new;User Id=Ffusr;Password=keR19a9;";
        static void Main(string[] args)
        {
            Welcome();
            readFile();
        }
        static void Welcome()
        {
            Console.WriteLine("Welcome to Database Restore Application.");
            Console.WriteLine("Press Any Key to Continue. ");
            Console.ReadKey();
        }
        static void readFile()
        {
            string[] listFiles;

            Console.WriteLine("Please Insert the Data Folder Path");
            folderName = "C:\\Users\\User\\Desktop\\TimeLogger";
            Console.Clear();
            listFiles = Directory.GetFiles(folderName);
            readExcelData(listFiles);
            Console.WriteLine("Process done with " + failed + " process.");
            Console.ReadKey();
        }
        static void readExcelData(string[] listFiles)
        {
            string deviceUserId,tarikhIn,tarikhOut,timeIn,timeOut;
            string terminalin, terminalout;
            char[] UID;
            int DeviceID;
            string updatedstring;
            //MySqlConnection conn = new MySqlConnection(sqlstring);
            SqlConnection msconn = new SqlConnection(mssqlstring);
            foreach (string s in listFiles)
            {
                try
                {
                    Random random = new Random();

                    //conn.Open();
                    msconn.Open();
                    Excel.Application excel = new Excel.Application();
                    Excel.Workbook excelbook = excel.Workbooks.Open(s);
                    Excel._Worksheet excelsheet = excelbook.Sheets[1];
                    Excel.Range range = excelsheet.UsedRange;
                    for (int i = 2; i <= range.Rows.Count; i++)
                    {
                        TIDIn = 0; TIDOut = 0;
                        deviceUserId = Convert.ToString(range.Cells[i, 2].Value2);
                        tarikhIn = range.Cells[i, 3].Text;
                        tarikhOut = range.Cells[i, 4].Text;
                        timeIn = range.Cells[i, 5].Text;
                        timeOut = range.Cells[i, 6].Text;
                        terminalin = range.Cells[i, 7].Text;
                        terminalout = range.Cells[i, 8].Text;
                        UID = deviceUserId.ToCharArray();
                        UID[0] = '3';
                        updatedstring = new string(UID);
                        if (timeIn != "NULL")
                        {
                            if (terminalin == "Terminal I          ")
                            {
                                TIDIn = 54;
                            }
                            else if (terminalin == "Terminal II         ")
                            {
                                TIDIn = 55;
                            }
                            else if (terminalin == "Terminal III        ")
                            {
                                TIDIn = 56;
                            }
                            else
                            {
                                TIDIn = 54;
                            }
                            Console.WriteLine(TIDIn + " " + tarikhIn + " " + timeIn);
                            commandmssqlIn = "INSERT INTO dbo.Raw (StaffID,TID,TimeIN,TimeID,LogType,FlagProses,EnrollID,LRawID,BranchID) VALUES(0," + TIDIn + ",'" + tarikhIn + " " + timeIn + "',0,0,0," + updatedstring + ",0,0)";
                            //commandIn = "INSERT INTO `raw`(`StaffID`, `TID`, `TimeIN`, `TimeID`, `LogType`, `FlagProses`, `EnrollID`, `FlagUpdate`) VALUES (0,'" + random.Next(6, 12).ToString() + "','" + tarikhIn + " " + timeIn + "',0,0,0," + updatedstring + ",0)";
                            //MySqlCommand command = new MySqlCommand(commandIn, conn);
                            SqlCommand commandms = new SqlCommand(commandmssqlIn, msconn);
                            try
                            {
                                //command.ExecuteNonQuery();
                                commandms.ExecuteNonQuery();
                            }
                            catch
                            {
                                failed++;
                            }
                        }
                        if (timeOut != "NULL")
                        {
                            if (terminalout == "Terminal I          ")
                            {
                                TIDOut = 54;
                            }
                            else if (terminalout == "Terminal II         ")
                            {
                                TIDOut = 55;
                            }
                            else if (terminalout == "Terminal III        ")
                            {
                                TIDOut = 56;
                            }
                            else
                            {
                                TIDOut = 54;
                            }
                            Console.WriteLine(tarikhOut + " " + timeOut);
                            commandmssqlOut = "INSERT INTO dbo.Raw (StaffID,TID,TimeIN,TimeID,LogType,FlagProses,EnrollID,LRawID,BranchID) VALUES(0," + TIDOut + ",'" + tarikhOut + " " + timeOut + "',0,0,0," + updatedstring + ",0,0)";
                            //commandOut = "INSERT INTO `raw`(`StaffID`, `TID`, `TimeIN`, `TimeID`, `LogType`, `FlagProses`, `EnrollID`, `FlagUpdate`) VALUES ('0','" + random.Next(6, 12).ToString() + "','" + tarikhOut + " " + timeOut + "','0','0','false','" + updatedstring + "','false')";
                            //MySqlCommand command = new MySqlCommand(commandOut, conn);
                            SqlCommand commandms = new SqlCommand(commandmssqlOut, msconn);
                            try
                            {
                                // command.ExecuteNonQuery();
                                commandms.ExecuteNonQuery();
                            }
                            catch
                            {
                                failed++;
                            }
                        }
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(excelsheet);

                    //close and release
                    excelbook.Close();
                    Marshal.ReleaseComObject(excelbook);

                    //quit and release
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);

                    //conn.Close();
                    msconn.Close();
                }
                catch
                {
                    continue;
                }
            }
        }
        public void Leave(string[] listFiles)
        {
            string deviceUserId, tarikhIn, tarikhOut, timeIn, timeOut;
            string terminalin, terminalout;
            char[] UID;
            int DeviceID;
            string updatedstring;
            //MySqlConnection conn = new MySqlConnection(sqlstring);
            SqlConnection msconn = new SqlConnection(mssqlstring);
            foreach (string s in listFiles)
            {
                try
                {
                    Random random = new Random();

                    //conn.Open();
                    msconn.Open();
                    Excel.Application excel = new Excel.Application();
                    Excel.Workbook excelbook = excel.Workbooks.Open(s);
                    Excel._Worksheet excelsheet = excelbook.Sheets[1];
                    Excel.Range range = excelsheet.UsedRange;
                    for (int i = 2; i <= range.Rows.Count; i++)
                    {
                        
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(excelsheet);

                    //close and release
                    excelbook.Close();
                    Marshal.ReleaseComObject(excelbook);

                    //quit and release
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);

                    //conn.Close();
                    msconn.Close();
                }
                catch
                {
                    continue;
                }
            }
        }
    }
}
