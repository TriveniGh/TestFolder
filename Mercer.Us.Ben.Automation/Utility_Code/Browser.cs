using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using Mercer.Us.Ben.Automation.Main;
using OpenQA.Selenium.Support.UI;
namespace Mercer.Us.Ben.Automation.Utility_Code
{
    public class Browser : Test
    {
        public static bool VerifyText(string PageName, string ObjectName, string Expval)
        {
            bool result = false;
            try
            {
                By OR_ObjectName = XML_Util.readObjectData(PageName, ObjectName);

                string ExpValue = Expval.Trim().ToLower();
                // IWebElement Obj = driver.FindElement(OR_ObjectName);
                IWebElement Obj = WaitfortheElementPresent(OR_ObjectName);

                string Obj_text = Obj.Text.Trim().ToLower();
                if (Obj_text.Equals(ExpValue))
                {

                    Logger.info(ObjectName + " Text verfied");
                    result = true;
                    Console.WriteLine("VerifyText" + result);
                    Test.Result(ObjectName + " should be displayed in" + PageName + "Page", ObjectName + "is displayed in" + PageName + "Page", "Passed");

                }
                else
                {
                    result = false;
                    Logger.info(" ");
                    Test.Result(ObjectName + "should be displayed in" + PageName + "Page", ObjectName + "is not displayed in" + PageName + "Page", "Failed");

                }

                return result;
            }
            catch (Exception e)
            {
                Logger.error("", e);
                result = false;
                Test.Result(ObjectName + "should be displayed in" + PageName + "Page", ObjectName + "is not displayed in" + PageName + "Page", "Failed");

                return result;
            }
            finally
            {
                Logger.info("In" + PageName + "*-_-*-_-*-_-*" + ObjectName);


            }
        }



        public static void IsPresent(string PageName, string ObjectName)
        {
            bool result = false;
            By OR_ObjectName = XML_Util.readObjectData(PageName, ObjectName);
            try
            {

                result = driver.FindElement(OR_ObjectName).Enabled;
                Console.WriteLine("IsPresent" + result);

                Test.Result(ObjectName + "should be displayed in" + PageName + "Page", ObjectName + "is displayed in" + PageName + "Page", "Passed");

            }
            catch (Exception e)
            {

                Logger.error("", e);
                Test.Result(ObjectName + "should be displayed in" + PageName + "Page", ObjectName + "is not displayed in" + PageName + "Page", "Failed");

            }
            finally
            {

                Logger.info("In" + PageName + "*-_-*-_-*-_-*" + ObjectName);
            }


        }

        public static IWebElement WaitfortheElementPresent(By by)
        {
            try
            {

                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(Constants.ElementPresent));
                return wait.Until(driv => driver.FindElement(by));

            }
            catch (Exception e)
            {


                Logger.error("Element does not found in the page", e);
                return null;

            }
            finally
            {
                Logger.info("Wait for the element Present *-_-*-_-*-_-*");


            }


        }

        public static bool IsDisplayed(string PageName, string ObjectName)
        {
            bool result = false;
            By OR_ObjectName = XML_Util.readObjectData(PageName, ObjectName);
            try
            {
               
                result = driver.FindElement(OR_ObjectName).Displayed;
                Console.WriteLine("IsDisplayed" + result);
                Test.Result(ObjectName + "should be displayed in" + PageName + "Page", ObjectName + "is displayed in" + PageName + "Page", "Passed");
                return result;
            }
            catch (Exception e)
            {

                Logger.error(ObjectName + "is not displayed in" + PageName + "Page", e);
                Test.Result(ObjectName + " should be displayed in" + PageName + "Page", ObjectName + "is not displayed in" + PageName + "Page", "Failed");
                Test.TakeScreenShot();
                return result;
            }
            finally
            {
                Logger.info("In" + PageName + "*-_-*-_-*-_-*" + ObjectName);

            }


        }

        public static void Click(string PageName, string ObjectName)
        {
            By OR_ObjectName = XML_Util.readObjectData(PageName, ObjectName);
            try
            {
                IWebElement Obj = WaitfortheElementPresent(OR_ObjectName);
                // driver.FindElement(OR_ObjectName).Click();
                if (Obj.Displayed)
                {
                    Obj.Click();
                }
                else
                {

                    ScrollDown(OR_ObjectName);

                    Obj.Click();

                }
            }
            catch (Exception e)
            {

                Logger.error(ObjectName + " is not found in " + PageName, e);

                Test.Result(ObjectName + " is not found in " + PageName, ObjectName + " is not found in " + PageName, "Failed");
                Test.TakeScreenShot();

            }
            finally
            {
                if (driver.Url.ToLower().Contains("error"))
                {
                    Test.Result("Logout should be successful", "Un expected error while performing logout", "Failed");
                    Test.TakeScreenShot();
                }
                Logger.info("In" + PageName + "*-_-*-_-*-_-*" + ObjectName);

            }

        }



        public static bool Launch(string URL)
        {
            try
            { 
                driver.Navigate().GoToUrl(URL);
                driver.Manage().Window.Maximize();

                if (driver.Title.ToLower().Contains("login") && !driver.Url.ToLower().Contains("error"))
                {
                    Test.Result("Application should Launch successfully", "Application  Launched successfully", "Passed");                    
                    return true;
                }
                else
                {
                    Test.Result("Application should Launch successfully", "Application  Failed to Launch", "Failed");
                    Constants.ApplicationLaunchFlag = false;
                    return false;
                 
                }
                //Console.WriteLine("*****************************************" + driver.Title);               
            }
            catch (Exception e)
            {
                Test.Result("Application should Launch successfully", "Application failed to Launch", "Failed");
                Logger.error("Application failed to Launch", e);
                return false;
            }
            finally
            {
                Logger.info("Launching the URL" + "*--*--*--*");
            }

        }




        public static void WaitforthePageLoad()
        {
            try
            {

              

               // driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(Constants.WaitforthePageLoad));

                IWait<IWebDriver> wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(Constants.WaitforthePageLoad));

        

                wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));

            }
            catch (Exception e)
            {
                Logger.error("Failed to load the Page", e);
                Test.Result("Page should be loaded", "Failed to Load the page", "Failed");
            }
            finally
            {
                Logger.info("  loading  the page" + "*--*--*--*");
                if (driver.Url.ToLower().Contains("error"))
                {
                    Test.Result("Logout should be successful", "Un expected error while performing logout", "Failed");
                    Test.TakeScreenShot();
                }
               // Logger.info("In" + PageName + "*-_-*-_-*-_-*" + ObjectName);
            }


        }
        public static void ScrollDown(By ObjectName)
        {
            try
            {
                var elem = driver.FindElement(ObjectName);
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", elem);

            }
            catch (Exception e)
            {
                Logger.error("Failed to load the Page", e);
                Test.Result("Page should be loaded", "Failed to Load the page", "Failed");
            }
            finally
            {
                Logger.info("Scrolling down the Page" + "*--*--*--*");
            }


        }


        public static void EnterText(string PageName, string ObjectName, string TextBoxtext)
        {
            By OR_ObjectName = XML_Util.readObjectData(PageName, ObjectName);
            try
            {

                driver.FindElement(OR_ObjectName).SendKeys(TextBoxtext);



            }
            catch (Exception e)
            {
                Logger.error("", e);
                Test.Result(ObjectName + "should be displayed in" + PageName + "Page", ObjectName + "is not  displayed in" + PageName + "Page", "Failed");

                Test.TakeScreenShot();

            }
            finally
            {
                Logger.info("In " + PageName + "*-_-*-_-*-_-*" + ObjectName);

            }

        }

        public static void Selectdropdown(string PageName, string ObjectName, string DropdownListText)  
        {
            By OR_ObjectName = XML_Util.readObjectData(PageName, ObjectName);
            try
            {
                SelectElement select = new SelectElement(driver.FindElement(OR_ObjectName));

                foreach (IWebElement element in select.Options)
                {
                    if (element.Text == DropdownListText)
                    {
                        element.Click();
                    }
                }

            }
            catch (Exception e)
            {
                Logger.error("", e);
                Test.Result(ObjectName + "should be displayed in" + PageName + "Page", ObjectName + "is not  displayed in" + PageName + "Page", "Failed");

                Test.TakeScreenShot();

            }
            finally
            {
                Logger.info("In " + PageName + "*-_-*-_-*-_-*" + ObjectName);

            }


        }

        //UnderProgress
        public static void ValidateDropdownoptions(string PageName, string ObjectName, string DropdownListText) 
        {
            By OR_ObjectName = XML_Util.readObjectData(PageName, ObjectName);
            try
            {
              IWebElement element = driver.FindElement(OR_ObjectName);
                IList<IWebElement> AllDropDownList = element.FindElements(OR_ObjectName);
                int DpListCount = AllDropDownList.Count;


            }
            catch (Exception e)
            {
                Logger.error("", e);
                Test.Result(ObjectName + "should be displayed in" + PageName + "Page", ObjectName + "is not  displayed in" + PageName + "Page", "Failed");

                Test.TakeScreenShot();

            }
            finally
            {
                Logger.info("In " + PageName + "*-_-*-_-*-_-*" + ObjectName);

            }


        }


        public static void VerifyPageTitleContains(string PageName, string PageTitle)
        {

            try
            {
                string Title = driver.Title;
                if (Title.Contains(PageTitle))
                {

                    Test.Result(PageTitle + " should be displayed in " + PageName, PageTitle + " is   displayed in " + PageName + " Page", "Passed");


                }
                else
                {


                    Test.Result(PageTitle + " should be displayed in " + PageName, PageTitle + " is not  displayed in " + PageName + " Page", "Failed");


                }



            }
            catch (Exception e)
            {
                Logger.error("", e);
                Test.Result(PageTitle + " should be displayed in " + PageName, PageTitle + " is not  displayed in " + PageName + " Page", "Failed");

                Test.TakeScreenShot();

            }
            finally
            {
                Logger.info("In " + PageName + "*-_-*-_-*-_-*" + PageTitle);

            }


        }
        public static void VerifyPageURLContains(string PageName, string PageURLContains)
        {

            try
            {
                string URL = driver.Url;
                if (URL.Contains(PageURLContains))
                {

                    Test.Result(PageURLContains + " should be displayed in " + PageName, PageURLContains + " is   displayed in " + PageName + " Page", "Passed");


                }
                else
                {


                    Test.Result(PageURLContains + " should be displayed in " + PageName, PageURLContains + " is not  displayed in " + PageName + " Page", "Failed");


                }



            }
            catch (Exception e)
            {
                Logger.error("", e);
                Test.Result(PageURLContains + " should be displayed in " + PageName, PageURLContains + " is not  displayed in " + PageName + " Page", "Failed");

                Test.TakeScreenShot();

            }
            finally
            {
                Logger.info("Verifying Page URL in  " + PageName + "*-_-*-_-*-_-*");

            }


        }
        public static void SwitchToNewWindow(string PageName)
        {
            try
            {
                driver.SwitchTo().Window(driver.WindowHandles.Last());

                driver.Manage().Window.Maximize();
                WaitforthePageLoad();
            }
            catch (Exception e)
            {
                Logger.error(" ", e);
            }
            finally
            {

            }
        }

                public static void SwitchToLastWindow(string PageName) {
            try
            {
                 //driver.Close();
                driver.SwitchTo().Window(driver.WindowHandles.First());

                driver.Manage().Window.Maximize();
                WaitforthePageLoad();
            }
            catch (Exception e)
            {
                Logger.error(" ", e);
            }
            finally { 
            
            }


        }

                public static void JavsScriptClick(string PageName, string ObjectName)
                {


                    By OR_ObjectName = XML_Util.readObjectData(PageName, ObjectName);
                    try
                    {
                        IWebElement element = driver.FindElement(OR_ObjectName);
                        IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
                        executor.ExecuteScript("arguments[0].click();", element);
                        Test.TakeScreenShot();

                    }
                    catch (Exception e)
                    {
                        Logger.error("", e);
                        Test.Result(ObjectName + " should be displayed in" + PageName + "Page", ObjectName + " is not  displayed in" + PageName + "Page", "Failed");

                      //  Test.TakeScreenShot();

                    }
                    finally
                    {
                        Test.TakeScreenShot();
                        if (driver.Url.ToLower().Contains("error"))
                        {
                            Test.Result("Logout should be successful", "Un expected error while performing logout", "Failed");
                            Test.TakeScreenShot();
                        }
                        Logger.info("In " + PageName + "*-_-*-_-*-_-*" + ObjectName);

                    }




                }
    }


}
