using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using System.IO;
using Aspose.Cells;
using System.Reflection;
using System.Linq;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing.Imaging;
using Mercer.Us.Ben.Automation.Main;
using Mercer.Us.Ben.Automation.Utility_Code;
using System.Xml.Linq;
using OpenQA.Selenium.Firefox;
using Mercer.Us.Ben.Automation.Repository;
using OpenQA.Selenium.Remote;


namespace Mercer.Us.Ben.Automation.Main
{
    public class Test
    {
        // Defining dictionary for storing Test Data and Method Name
        // Defining dictionary for storing Test Data and Method Name
        private static Dictionary<string, string> values = new Dictionary<string, string>();
        static string fileName = "";
        public static string OutputfilePath = "";
        public  static string CurrentProjectPath = "";
        public  static string Sheet_Name = "";
        static string Scinario_Name = "";
        static string TestDescription = "";
        static string FunctionName = "";
        static string ResourcePath = "";
        public static IWebDriver driver;
        static string Browser_Type;
        static string BrowserDriverPath;
        int Row_count;
        static string Starttime = "";
        static int Passed = 0;
        static int Failed = 0;
        static string ActualStartTime = "";
        static string BrowserName = "";
        public static string strScreeShotPath;

        public void RunTest(String TestType)
        {
            try
            {

                Logger.createLogFile();
                ActualStartTime = DateTime.Now.ToString("hh:mm:ss tt", System.Globalization.DateTimeFormatInfo.InvariantInfo);
                CurrentProjectPath = ResultSheet.CurrentProjectPath();
                OutputfilePath = ResultSheet.ResultFilePath_Name();
                Email_Outlook.sendEMailThroughOUTLOOK();
                ResultSheet.CreateExcelReport(OutputfilePath);
                BrowserDriverPath = CurrentProjectPath + "Mercer.US.Ben.Automation" + Path.DirectorySeparatorChar + "BrowserDrivers";

                switch (TestType)
                {
                    case "Smoke":
                        ResourcePath = CurrentProjectPath + "Mercer.US.Ben.Automation" + Path.DirectorySeparatorChar + "Resources" + Path.DirectorySeparatorChar + "Smoke";
                        Logger.info("Framing Smoke test Rresource path :" + ResourcePath);
                        break;
                    case "Regression":
                        ResourcePath = CurrentProjectPath + "Mercer.US.Ben.Automation" + Path.DirectorySeparatorChar + "Resources" + Path.DirectorySeparatorChar + "Regression";
                        Logger.info("Framing Regression test Resource path :" + ResourcePath);
                        break;
                }
                fileName = ResourcePath + Path.DirectorySeparatorChar + "MODULES.xlsx";
                Logger.info("Module sheet path is :" + fileName);

                //  ModuleList for loading various module names 
                List<string> ModuleList = new List<string>();
                List<string> BrowserType = new List<string>();
                {
                    //open Modules Excel sheet 
                    Logger.info("Following Excel sheet loaded successfuly " + fileName);
                    FileStream MasterExcel = new FileStream(fileName, FileMode.Open);
                    Aspose.Cells.LoadOptions loadOptions = new Aspose.Cells.LoadOptions(LoadFormat.Xlsx);
                    Workbook wb = new Workbook(MasterExcel, loadOptions);
                    Worksheet ws = wb.Worksheets[0];

                    // Number of Rows in Module Sheet
                    Row_count = ws.Cells.Rows.Count;
                    Logger.info("Total number of Modules listed in the modul execution sheet" + Row_count);

                    for (int i = 1; (i <= Row_count); i++)
                    {
                        if (((ws.Cells[i, 0]).StringValue).Equals("Y"))
                        {
                            //Getting Module names 
                            string strModuleName = (ws.Cells[i, 1]).StringValue;
                            //strBrowserName = ;
                            BrowserType.Add(((ws.Cells[i, 2]).StringValue));
                            Logger.info("Selected Browser name is : " + ((ws.Cells[i, 2]).StringValue));
                            Logger.info("Currently loadding module is : " + strModuleName);
                            ModuleList.Add(strModuleName);
                        }
                    }
                }



                //For loop for Executing Various selected Modules
                for (int h = 0; h <= (ModuleList.Count) - 1; h++)
                {
                    Sheet_Name = ModuleList[h];
                    string strCurrentDataSheetName = ResourcePath + Path.DirectorySeparatorChar + ModuleList[h] + ".xlsx";
                    Logger.info("The following test data sheet loaded successfuly" + strCurrentDataSheetName);

                    List<string> Scenario_List = new List<string>();
                    List<string> ScenarioDescription_List = new List<string>();
                    {
                        FileStream ModuleExcel = new FileStream(strCurrentDataSheetName, FileMode.Open);
                        Aspose.Cells.LoadOptions loadOptions_1 = new Aspose.Cells.LoadOptions(LoadFormat.Xlsx);
                        Workbook wb_1 = new Workbook(ModuleExcel, loadOptions_1);
                        Worksheet ws_1 = wb_1.Worksheets[0];

                        int Row_count_1 = ws_1.Cells.Rows.Count;
                        Logger.info(" Total number of scenarios to be executed as per Test Data sheet" + Row_count_1);



                        //Loop for getting Scenario Names with Y 
                        for (int a = 1; (a <= Row_count_1); a++)
                        {
                            if (((ws_1.Cells[a, 0]).StringValue).Equals("Y"))
                            {
                                Scenario_List.Add((ws_1.Cells[a, 1]).StringValue);
                                Logger.info("Current executable Scenario name :" + ((ws_1.Cells[a, 1]).StringValue));

                                ScenarioDescription_List.Add((ws_1.Cells[a, 2]).StringValue);
                                Logger.info("The Scenario Description is:" + ((ws_1.Cells[a, 2]).StringValue));
                            }
                        }

                        // for loop for  Iterating over Various Scenarios
                        for (int s = 0; s <= (Scenario_List.Count) - 1; s++)
                        {
                            String value = Scenario_List[s];
                            Scinario_Name = ScenarioDescription_List[s];
                            Logger.info("Navigating to sheet and taking function name into Function_Name_List  : " + value);

                            // start the Browser invoke code
                            //For opening Browser
                            Logger.info("Launching Browser window");

                            //Get the browser type to be executed
                            //browserconfig > Object containing browsersettings.config file values
                            XDocument browserconfig = XDocument.Load(CurrentProjectPath + "Mercer.Us.Ben.Automation" + Path.DirectorySeparatorChar + "BrowserSettings.config");
                            Browser_Type = browserconfig.Root.Element("BrowserSelection").Element("Browser").Attribute("Type").Value;

                            if (Browser_Type == null || Browser_Type == "")
                            {
                                Browser_Type = "InternetExplorer";
                                Logger.info("Default Browser Setting is not set in BrowserSettingConfig file so its taking Firefox as default browser");
                            }

                            if (BrowserType[h] != null || BrowserType[h] != "")
                            {
                                Browser_Type = BrowserType[h];
                            }

                            Logger.info("Current Module is execuing agains Browser as :" + Browser_Type);
                            switch (Browser_Type)
                            {
                                case "Firefox":

                                    BrowserName = "Firefox";
                                    //Attempt to start Firefox. Tries consumer Firefox first, then attempts ESR version. 
                                    FirefoxBinary binary = new FirefoxBinary(browserconfig.Root.Element("BrowserDriver").Element("Firefox").Attribute("Path").Value);
                                    FirefoxProfile profile = new FirefoxProfile();

                                    driver = new FirefoxDriver(binary, profile);
                                    ICapabilities cap = ((RemoteWebDriver)driver).Capabilities;
                                    String browserName = cap.BrowserName.ToString();
                                    Console.WriteLine("*************" + browserName);
                                    string version = cap.Version.ToString();
                                    Console.WriteLine("*************" + version);

                                    Logger.info("FireFox driver is loaded");
                                    break;

                                case "InternetExplorer":
                                    //driver = new InternetExplorerDriver(BrowserDriverPath);
                                    //Logger.info("IE driver is selected");
                                    BrowserName = "IE";
                                    var options = new InternetExplorerOptions()
                                    {
                                        EnsureCleanSession = Convert.ToBoolean(browserconfig.Root.Element("BrowserDriver")
                                            .Element("InternetExplorer").Attribute("EnsureCleanSession").Value),
                                        InitialBrowserUrl = browserconfig.Root.Element("BrowserDriver")
                                            .Element("InternetExplorer").Attribute("InitialBrowserUrl").Value,
                                        IntroduceInstabilityByIgnoringProtectedModeSettings = Convert.ToBoolean(browserconfig.Root
                                            .Element("BrowserDriver").Element("InternetExplorer").Attribute("IgnoreProtectedMode").Value)
                                    };
                                    //Optional for Driver path---> browserconfig.Root.Element("BrowserDriver").Element("InternetExplorer").Attribute("Path").Value
                                    driver = new InternetExplorerDriver(BrowserDriverPath, options);
                                    Logger.info("IE driver is loaded");
                                    break;

                                case "Chrome":
                                    BrowserName = "Chrome";
                                    Logger.info("Chrome driver is selected");
                                    break;

                            }

                            //Navigating to sheet and taking function name into Function_Name_List
                            Worksheet worksheet = wb_1.Worksheets[value];
                            {
                                int Row_count_2 = worksheet.Cells.Rows.Count;
                                Logger.info("Getting row count of Test methods to be executed sheet: " + Row_count_2);

                                int col_Count_2 = worksheet.Cells.Columns.Count;
                                Logger.info("Getting column count of Test methods to be executed sheet: " + col_Count_2);

                                for (int j = 1; (j <= Row_count_2); j++)
                                {
                                    if (((worksheet.Cells[j, 0]).StringValue).Equals("Y"))
                                    {
                                        //Getting functionm names for getting test data
                                        //for Clearing the values of Dictionary for Each iteration 
                                        values.Clear();
                                        FunctionName = (worksheet.Cells[j, 2]).StringValue;
                                        values.Add("Function_Name_Value", FunctionName);
                                        Logger.info("Current executing function name is : " + FunctionName);
                                        string ClassName = (worksheet.Cells[j, 1]).StringValue;
                                        values.Add("Class_Name", ClassName);
                                        Logger.info("Current executing class name is : " + ClassName);
                                        TestDescription = (worksheet.Cells[j, 3]).StringValue;
                                        Logger.info("Test scenario description is : " + TestDescription);

                                        for (int k = 7; k <= col_Count_2; k++)
                                        {
                                            string Test_Value = (worksheet.Cells[j, k]).StringValue;
                                            string Test_Data = (worksheet.Cells[j + 1, k]).StringValue;

                                            if (Test_Value == null || Test_Value == "")
                                            {
                                                Logger.info("Getting out from the scenario test data collection list");
                                                break;
                                            }
                                            values.Add(Test_Value, Test_Data);
                                            Logger.info("The Parameter Name is : " + Test_Value + " and value is : " + Test_Data);
                                            string Exit_Test_Value = (worksheet.Cells[j, k + 1]).StringValue;
                                            if (Exit_Test_Value == null || Exit_Test_Value == "")
                                            {
                                                Logger.info("Getting out from the scenario execution list");
                                                break;
                                            }
                                        }
                                        // end Browser invoke code

                                        string ClassNameValue = values["Class_Name"];
                                        string MethodName = values["Function_Name_Value"];

                                        Logger.info("the name of the class  : " + ClassNameValue);
                                        Logger.info("the name of the Function or Method Name  : " + MethodName);

                                        string TypeName = "Mercer.Us.Ben.Automation.Test_Methods.MBC." + ClassNameValue;
                                        Type type = Type.GetType(TypeName);
                                        //    ConstructorInfo ctor = type.GetConstructor(new[] { typeof(IWebDriver) });
                                        //    object instance = ctor.Invoke(new object[] { driver });
                                        object classInstance = Activator.CreateInstance(type, null);
                                        MethodInfo methodInfo = type.GetMethod(MethodName);
                                        ParameterInfo[] parameters = methodInfo.GetParameters();

                                        //Navigate to Method  //Reflection 
                                        //result = (string)methodInfo.Invoke(instance, null);
                                        // methodInfo.Invoke(instance, null);
                                        Starttime = DateTime.Now.ToString("hh:mm:ss tt", System.Globalization.DateTimeFormatInfo.InvariantInfo);

                                        methodInfo.Invoke(classInstance, null);
                                        if (Constants.ApplicationLaunchFlag == false)
                                        {
                                            driver.Close();
                                            break;
                                        }
                                        Logger.info(" Status of the current execution");
                                    }
                                    j++;//loop increment for row
                                }
                            }
                        }
                        driver.Quit();
                    }

                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Test.Result("Page should be loaded", " Waiting for page load/Failed to page load", "Failed");
                Logger.error(e);
            }
            finally
            {
                ResultSheet.DesignChart(OutputfilePath, Passed, Failed, ActualStartTime, BrowserName);
                Email_Outlook.send_Report_EMailThroughOUTLOOK();
                Console.WriteLine("Test Execution is Completed");
            }
        }

        public static void Result(string Expected_Result, string Actual_Result, string Comments)
        {

            Microsoft.Office.Interop.Excel.Application objExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook objworkbook;
            Microsoft.Office.Interop.Excel.Worksheet objworksheet;

            try
            {
                string EndTime = DateTime.Now.ToString("hh:mm:ss tt", System.Globalization.DateTimeFormatInfo.InvariantInfo);
                objExcel.DisplayAlerts = false;
                objExcel.Visible = false;
                Excel.Range range = null;

                objworkbook = objExcel.Workbooks.Open(OutputfilePath, true, false,
                                                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                      Type.Missing, Type.Missing, true, Type.Missing,
                                                      Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                objworksheet = (Microsoft.Office.Interop.Excel.Worksheet)objworkbook.Sheets.get_Item("Results");
                TakeScreenShot();
                range = objworksheet.UsedRange;

                int RowCount = range.Rows.Count;
                int G = RowCount + 1;

                objworksheet.Cells[G, 1] = RowCount;
                objworksheet.Cells[G, 3] = BrowserName;
                objworksheet.Cells[G, 4] = Starttime;
                objworksheet.Cells[G, 5] = EndTime;
                objworksheet.Cells[G, 6] = Sheet_Name;

                objworksheet.Cells[G, 7] = Scinario_Name;
                objworksheet.Cells[G, 8] = TestDescription;
                objworksheet.Cells[G, 9] = FunctionName;
                objworksheet.Cells[G, 10] = Expected_Result;
                objworksheet.Cells[G, 11] = Actual_Result;
                objworksheet.Cells[G, 12] = strScreeShotPath;
                objworksheet.Hyperlinks.Add(objworksheet.Cells[G, 12], strScreeShotPath, Type.Missing, Type.Missing, "Snapshot");

                if (Comments == "Passed" || Comments == "PASSED" || Comments == "passed")
                {
                    Passed = Passed + 1;
                    range = objworksheet.get_Range("b" + G);
                    objworksheet.Cells[G, 2] = Comments;

                    range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                    range.Font.Size = 10;



                }
                else if (Comments == "Failed" || Comments == "failed" || Comments == "FAILED")
                {
                    Failed = Failed + 1;
                    range = objworksheet.get_Range("a" + G, "k" + G);
                    objworksheet.Cells[G, 2] = Comments;

                    range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    range.Font.Size = 10;




                }
                range = objworksheet.UsedRange;
                range.Columns.AutoFit();
                range.BorderAround(Excel.XlLineStyle.xlContinuous,
                Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic,
                Excel.XlColorIndex.xlColorIndexAutomatic);




                objworksheet.Name = "Results";
                object objOpt = Missing.Value;
                objworkbook.Save();
                objworkbook.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
                objExcel.Quit();

            }
            catch (Exception ex)
            {
                throw ex;
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
        public static string Data(string word)
        {
            // Try to get the result in the static Dictionary
            string result;
            if (values.TryGetValue(word, out result))
            {
                return result;
            }
            else
            {
                return null;
            }
        }

        public static void TakeScreenShot()
        {


            DateTime cal = DateTime.Now;
            string date_time = string.Format("{0:MM_dd_yyyy_HH_mm_ss}", cal);

            //string Resultfolder = solutionPath + "Mercer.US.Ben.Automation" + Path.DirectorySeparatorChar + "Results" + Path.DirectorySeparatorChar + datevalue + Path.DirectorySeparatorChar;

            string projectName = ".xls";
            string screenshotPath = OutputfilePath.Replace(projectName, "");
            string screenshotfolder = screenshotPath + Path.DirectorySeparatorChar;

            if (!Directory.Exists(screenshotfolder))
            {
                Directory.CreateDirectory(screenshotfolder);
                Logger.info("Results folder path is created!!");
            }
            else
            {
                Logger.info("Results folder path is already created!!");
            }



            Rectangle bounds = Screen.GetBounds(Point.Empty);
            using (Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {


                    g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                }

                strScreeShotPath = screenshotfolder + FunctionName + date_time + ".jpg";
                bitmap.Save(strScreeShotPath, ImageFormat.Jpeg);
            }

        }

    }
}
