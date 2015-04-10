using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;
using System.Net;
using CsvHelper;
using System.IO;

namespace UrlToFile {
    class Program {

        private static string _sourceFile = String.Empty;
        private static string _outputFormat = @"c:\temp\{0}_{1}.jpg";               // Russ Reid
        //private static string _outputFormat = @"c:\temp\Child_{0}-{1}.jpg";        // Blue North

        static void Main(string[] args) {

            if (args.Length == 0) {
                Usage();
                return;
            }

            // gather all the parameters
            foreach (string arg in args) {
                if (arg.StartsWith("/S:", StringComparison.InvariantCultureIgnoreCase)) {
                    _sourceFile = arg.Substring(3);
                }
                if (arg.StartsWith("/O:", StringComparison.InvariantCultureIgnoreCase)) {
                    _outputFormat = arg.Substring(3);
                }
            }

            // check existence of source file
            if (_sourceFile == string.Empty) {
                Usage();
                return;
            }

            #region Excel Format
            
            //Console.WriteLine("source is [" + sourceFile + "]");
            //// open the source file
            //Application xlApp = new Application();
            //Workbook xlWorkbook = null;
            //Worksheet xlWorksheet = null;
            //Range range;

            //try {
            //    xlWorkbook = xlApp.Workbooks.Open(
            //            sourceFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //    xlWorksheet = xlApp.ActiveWorkbook.ActiveSheet;


            //    range = xlWorksheet.UsedRange;

            //    string str;
            //    int rCnt = 0;
            //    int cCnt = 0;

            //    for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++) {
            //        for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++) {
            //            str = (string)(range.Cells[rCnt, cCnt] as Range).Value2;
            //            Console.WriteLine(string.Format("[{0},{1}]:[{2}]", rCnt, cCnt, str));
            //        }
            //    }

            //    xlWorkbook.Close(true, null, null);
            //    xlApp.Quit();

            //}
            //catch (Exception ex) {

            //    Console.WriteLine("error encountered: " + ex.Message);
            //}
            //finally {
            //    ReleaseObject(xlWorksheet);
            //    ReleaseObject(xlWorkbook);
            //    ReleaseObject(xlApp);

            //}
            #endregion

            #region Excel into Dataset
            
            // =======================================


            //OleDbConnection MyConnection = null;
            //WebClient webClient = new WebClient();
            //try {
            //    System.Data.DataSet DtSet;
            //    System.Data.OleDb.OleDbDataAdapter MyCommand;

            //    string connString = string.Format(@"provider=Microsoft.Jet.OLEDB.4.0;Data Source='{0}';Extended Properties=Excel 8.0;", sourceFile);

            //    MyConnection = new OleDbConnection(connString);
            //    MyCommand = new OleDbDataAdapter("select * from [Sheet1$]", MyConnection);
            //    MyCommand.TableMappings.Add("Table", "TestTable");
            //    DtSet = new System.Data.DataSet();
            //    MyCommand.Fill(DtSet);

            //    int rowCount = 0;
            //    foreach (DataRow row in DtSet.Tables[0].Rows) {
            //        string childId = (string)row["child"];
            //        string url = (string)row["url"];

            //        string target = string.Format(outputFormat, childId);
            //        try {
            //            webClient.DownloadFile(url, target);
            //        }
            //        catch (Exception ex) {
            //            // todo: add this to error file.
            //            Console.WriteLine(string.Format("error row {4}: copying url [{1}] to [{2}]{0}{3}",
            //                    Environment.NewLine, url, target, ex.Message, rowCount));
            //        }

            //        rowCount++;
            //    }
            //    //dataGridView1.DataSource = DtSet.Tables[0];
            //}
            //catch (Exception ex) {
            //    Console.WriteLine("error encountered: " + ex.Message);
            //}
            //finally {
            //    webClient.Dispose();
            //    MyConnection.Close();
            //}
            #endregion

            // -------------------------

            Console.WriteLine("UrlToFile Utility. World Vision Canada. @2015 NMTT");
            Console.WriteLine();
            Console.WriteLine("Source file: [{0}]", _sourceFile);
            Console.WriteLine("Output format: [{0}]", _outputFormat);

            WebClient webClient = new WebClient();

            try {
                using (TextReader reader = File.OpenText(_sourceFile)) {
                    var csv = new CsvReader(reader);
                    csv.Configuration.HasHeaderRecord = false;

                    int rowCount = 0;
                    while (csv.Read()) {
                        Console.WriteLine("Row:{0}...", rowCount);

                        string country = csv.GetField(0);
                        string childId = csv.GetField(1);
                        string url = csv.GetField(2);

                        string target = string.Format(_outputFormat, country, childId);

                        Console.WriteLine("   url: [{0}]", url);
                        Console.WriteLine("  dest: [{0}]", target);

                        try {
                            webClient.DownloadFile(url, target);
                        }
                        catch (Exception ex) {
                            // todo: add this to error file.
                            Console.WriteLine(string.Format("error row {4}: copying url [{1}] to [{2}]{0}{3}",
                                    Environment.NewLine, url, target, ex.Message, rowCount));
                        }
                        rowCount++;
                    }
                }

            }
            catch (Exception ex) {
                Console.WriteLine("error encountered: " + ex.Message);
            }
            finally {
                webClient.Dispose();
            }

            Console.Write("Done. Press <Enter> to exit...");
            Console.ReadLine();
        }

        private static void Usage() {
            Console.WriteLine("Converts URLs into actual image files.");
            Console.WriteLine();
            Console.WriteLine("UrlToFile /S:[drive:][path]filename [/O:[drive:][path]fileformat]");
            Console.WriteLine();
            Console.WriteLine("/S:[drive:][path]filename        Comma Separated Value (CSV) source file (no header) with the record format of: <CounrtyCode>,<ChildId>,<ImageUrl>");
            Console.WriteLine("/O:[drive:][path]fileformat      Output fileformat {0} corresponds to the first column of the csv file (country); {1} corresponds to the second column of the csv file (child id).");

            Console.ReadLine();
        }

        private static void ReleaseObject(object obj) {
            try {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex) {
                obj = null;
                Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally {
                GC.Collect();
            }
        }
    }
}
