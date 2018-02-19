using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace Mercer.Us.Ben.Automation.Utility_Code
{
  public  class ResultSheet
    {
        public static string CurrentProjectPath()
        {
            string currPrjDirpath = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory()));
            string projectName = "TestResults";
            string solutionPath = currPrjDirpath.Replace(projectName, "");

            return solutionPath;
        }

        public static string ResultFilePath_Name()
        {

            DateTime date = DateTime.Now;
            string datevalue = string.Format("{0:yyyy_MM_dd}", date);
            DateTime cal = DateTime.Now;
            string date_time = string.Format("{0:MM_dd_yyyy_HH_mm_ss}", cal);

            string currPrjDirpath = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory()));
            string projectName = "TestResults";
            string solutionPath = currPrjDirpath.Replace(projectName, "");
            string Resultfolder = solutionPath + "Mercer.US.Ben.Automation" + Path.DirectorySeparatorChar + "Results" + Path.DirectorySeparatorChar + datevalue + Path.DirectorySeparatorChar;
            if (!Directory.Exists(Resultfolder))
            {
                Directory.CreateDirectory(Resultfolder);
                Console.WriteLine("Results folder path is created!!");
            }
            else
            {
                Console.WriteLine("Results folder path is already created!!");
            }

            string Excelfilepath = Resultfolder + date_time + ".xls";


            return Excelfilepath;
        }

        public static string CreateExcelReport(string Excelfilepath)
        {
            if (!File.Exists(Excelfilepath))
            {
                string[] strArray = new string[] { "Sl.No", "Comments", "BrowserName", "StartTime", "EndTime", "ModuleName", "Scenario Description", "Test Description", "MethodName", "Expected Result", "Actual Result" };

                int Array_Length = strArray.Length;
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Range formatRange;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                for (int Q = 0; Q <= (Array_Length - 1); Q++)
                {
                    xlWorkSheet.Cells[1, Q + 1] = strArray[Q];

                    xlWorkSheet.Cells[1, Q + 1].Font.Bold = true;
                }

                xlWorkSheet.Name = "Results";
                formatRange = xlWorkSheet.get_Range("a1", "k1");

                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                formatRange = xlWorkSheet.get_Range("a1", "k1");
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                formatRange.Font.Size = 10;

                formatRange.Columns.AutoFit();

                xlWorkBook.SaveAs(Excelfilepath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);


            }
            return Excelfilepath;
        }


        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                obj = null;

            }
            finally
            {
                GC.Collect();
            }
        }


        public static void DesignChart(string OutputfilePath, int Passed, int Failed, string ActualStartTime, string BrowserName)
        {
            Microsoft.Office.Interop.Excel.Application objExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook objworkbook;
            object misValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Worksheet objworksheet;

            try
            {


                string EndTime = DateTime.Now.ToString("hh:mm:ss tt", System.Globalization.DateTimeFormatInfo.InvariantInfo);
                objExcel.DisplayAlerts = false;
                objExcel.Visible = false;
                Excel.Range range = null;
                // int row = 1, col = 1;
                objworkbook = objExcel.Workbooks.Open(OutputfilePath, true, false,
                                                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                      Type.Missing, Type.Missing, true, Type.Missing,
                                                      Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                objworksheet = (Microsoft.Office.Interop.Excel.Worksheet)objworkbook.Sheets.get_Item("Results");
                range = objworksheet.UsedRange;




                objworksheet =
                         (Microsoft.Office.Interop.Excel.Worksheet)objworkbook.Sheets.Add();
                objworksheet = (Microsoft.Office.Interop.Excel.Worksheet)objworkbook.Sheets.get_Item("Sheet1");



                objworksheet.Cells[1, 1] = "";
                range = objworksheet.get_Range("b1");
                objworksheet.Cells[1, 2] = "Mercer.US.Benefits";
                objworksheet.Cells[1, 2].Font.Bold = true;
                //formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                range.Font.Size = 15;




                range = objworksheet.get_Range("h2");
                objworksheet.Cells[2, 8] = "Mercer.US.Benefits";
                objworksheet.Cells[2, 8].Font.Bold = true;
                range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                range.Font.Size = 15;


                range = objworksheet.get_Range("a2");
                objworksheet.Cells[2, 1] = "Passed";
                objworksheet.Cells[2, 1].Font.Bold = true;
                range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                range.Font.Size = 15;

                objworksheet.Cells[2, 2] = Passed;
                objworksheet.Cells[2, 2].Font.Bold = true;
                range = objworksheet.get_Range("b2");
                range.Font.Size = 15;


                range = objworksheet.get_Range("a3");
                range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                objworksheet.Cells[3, 1] = "Failed";
                objworksheet.Cells[3, 1].Font.Bold = true;
                range.Font.Size = 15;


                range = objworksheet.get_Range("b3");
                objworksheet.Cells[3, 2] = Failed;
                objworksheet.Cells[3, 2].Font.Bold = true;
                range = objworksheet.get_Range("b3");
                range.Font.Size = 15;

                range = objworksheet.get_Range("a4");
                objworksheet.Cells[4, 1] = "Skipped";
                objworksheet.Cells[4, 1].Font.Bold = true;
                range.Font.Size = 15;

                range = objworksheet.get_Range("l4");
                objworksheet.Cells[4, 12] = "StartTime" + "-->" + ActualStartTime;
                range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                objworksheet.Cells[4, 12].Font.Bold = true;
                range.Font.Size = 15;

                range = objworksheet.get_Range("l5");
                objworksheet.Cells[5, 12] = "EndTime" + "  " + "-->" + EndTime;
                range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                objworksheet.Cells[5, 12].Font.Bold = true;
                range.Font.Size = 15;

                range = objworksheet.get_Range("l6");
                objworksheet.Cells[6, 12] = "BrowserName" + "-->" + BrowserName;
                range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                objworksheet.Cells[6, 12].Font.Bold = true;
                range.Font.Size = 15;



                objworksheet.Cells[4, 2] = "";
                Excel.Range chartRange;
                Excel.ChartObjects xlCharts = (Excel.ChartObjects)objworksheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(100, 100, 300, 250);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = objworksheet.get_Range("A1", "d5");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.ChartType = Excel.XlChartType.xl3DPie;
                ((Excel.LegendEntry)chartPage.Legend.LegendEntries(2)).LegendKey.Interior.Color = (int)Excel.XlRgbColor.rgbRed;
                ((Excel.LegendEntry)chartPage.Legend.LegendEntries(1)).LegendKey.Interior.Color = (int)Excel.XlRgbColor.rgbGreen;
                ((Excel.LegendEntry)chartPage.Legend.LegendEntries(4)).LegendKey.Interior.Color = (int)Excel.XlRgbColor.rgbWhite;


                range = objworksheet.get_Range("a1", "j1");
                range.Font.Size = 10;

                range.Columns.AutoFit();




                objworksheet.Name = "Output";
                object objOpt = Missing.Value;
                objworkbook.Save();
                objworkbook.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
                objExcel.Quit();





            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                objExcel = null;
                objworkbook = null;
                objworksheet = null;
                ReleaseComObject(objExcel);
                ReleaseComObject(objworkbook);
                ReleaseComObject(objworksheet);
            }
        }


        public static void ReleaseComObject(object reference)
        {
            try
            {
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(reference) <= 0)
                {

                }
            }
            catch
            {
            }
        }

    }
}
