using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;
using System.Net;

namespace UrlToFile {
    class Program {

        private static string sourceFile = String.Empty;
        private static string outputFormat = @"c:\temp\{0}.jpg";
        private static string errorFile = @"c:\temp\UrlToFileError.txt";

        static void Main(string[] args) {

            if (args.Length == 0) {
                Usage();
                return;
            }

            // gather all the parameters
            foreach (string arg in args) {
                Console.WriteLine("arg: [" + arg + "]");

                if (arg.StartsWith("/S:", StringComparison.InvariantCultureIgnoreCase)) {
                    sourceFile = arg.Substring(3);
                }
                if (arg.StartsWith("/O:", StringComparison.InvariantCultureIgnoreCase)) {
                    outputFormat = arg.Substring(3);
                }
                if (arg.StartsWith("/E:", StringComparison.InvariantCultureIgnoreCase)) {
                    errorFile = arg.Substring(3);
                }
            }

            // check existence of source file
            if (sourceFile == string.Empty) {
                Usage();
                return;
            }

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




            OleDbConnection MyConnection = null;
            WebClient webClient = new WebClient();
            try {
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;

                string connString = string.Format(@"provider=Microsoft.Jet.OLEDB.4.0;Data Source='{0}';Extended Properties=Excel 8.0;", sourceFile);

                MyConnection = new OleDbConnection(connString);
                MyCommand = new OleDbDataAdapter("select * from [Sheet1$]", MyConnection);
                MyCommand.TableMappings.Add("Table", "TestTable");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);

                int rowCount = 0;
                foreach (DataRow row in DtSet.Tables[0].Rows) {
                    string childId = (string)row["child"];
                    string url = (string)row["url"];

                    string target = string.Format(outputFormat, childId);
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
                //dataGridView1.DataSource = DtSet.Tables[0];
            }
            catch (Exception ex) {
                Console.WriteLine("error encountered: " + ex.Message);
            }
            finally {
                webClient.Dispose();
                MyConnection.Close();
            }
            Console.Write("press <Enter> to exit...");
            Console.ReadLine();
        }

        private static void Usage() {
            Console.WriteLine("Converts URLs into actual image files.");
            Console.WriteLine();
            Console.WriteLine("UrlToFile /S:[drive:][path]filename [/O:[drive:][path]fileformat] [/E:[drive:][path]filename]");
            Console.WriteLine();
            Console.WriteLine("/S:[drive:][path]filename        Source file.");
            Console.WriteLine("/O:[drive:][path]fileformat      Output fileformat.");
            Console.WriteLine("/E:[drive:][path]filename        Error file.");

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
