using Mercer.Us.Ben.Automation.Main;
using Mercer.Us.Ben.Automation.Utility_Code;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mercer.Us.Ben.Automation.Test_Methods.MBC
{
    public class MBC_Proxy : Test
    {

        public void ImpersonateFunc()
        {
            Logger.info("Entering to 'ImpersonateFunc' method");
            try
            {
                if (Test.Data("URL") != "" || Test.Data("URL") != null)
                {
                    Browser.Launch(Test.Data("URL"));
                    Browser.WaitforthePageLoad();
                    Browser.EnterText("MBC_Proxy", "MBCUserName", Test.Data("UserName"));
                    Browser.EnterText("MBC_Proxy", "MBCPassword", Test.Data("Password"));
                    Browser.Click("MBC_Proxy", "LoginButton");
                    Browser.WaitforthePageLoad();

                    Browser.Click("MBC_Proxy", "AdminAccessLink");
                    Browser.WaitforthePageLoad();
                    Browser.IsDisplayed("EmployeeSearchPage", "EmpSearchButton");
                    Browser.VerifyText("EmployeeSearchPage", "EmpSearchLabel", "Employee Search page for Proxy");
                    Browser.Click("EmployeeSearchPage", "EmpSearchButton");
                    Browser.WaitforthePageLoad();
                    Browser.VerifyText("WhatToDoHerePage", "EmployeeSearchLabel", "Employee Search");
                    Browser.VerifyText("WhatToDoHerePage", "WhatToDoLabel", "What To Do Here");
                    Browser.VerifyText("WhatToDoHerePage", "ToInitiateSearchLabel", "To initiate a search for an employee, enter all or part of the last name, employee number or the SSN of the employee you wish to find. Click Search to begin your search.");
                    Browser.VerifyText("WhatToDoHerePage", "SearchByLabel", "Search By");
                    Browser.IsDisplayed("WhatToDoHerePage", "PrintButton");
                    Browser.VerifyText("WhatToDoHerePage", "KeywordLabel", "Keyword:");
                    Browser.VerifyText("WhatToDoHerePage", "Tips", "Tip: Use % as a wildcard, for example, searching for Jam% will produce results for James and Jamison");
                    if ((Test.Data("Last Name") == null) || (Test.Data("Last Name") == ""))
                    {
                        Logger.info("LastName Field is empty");
                    }
                    else
                    {
                        Browser.Selectdropdown("WhatToDoHerePage", "SearchByDropDown", "Last Name");
                        Browser.EnterText("WhatToDoHerePage", "KeywordTextBox", Test.Data("LastName"));
                    }
                    if ((Test.Data("Employee Number") == null) || (Test.Data("Employee Number") == ""))
                    {
                        Logger.info("Employee Number Field is empty");
                    }
                    else
                    {
                        Browser.Selectdropdown("WhatToDoHerePage", "SearchByDropDown", "Employee Number");
                        Browser.EnterText("WhatToDoHerePage", "KeywordTextBox", Test.Data("Employee Number"));
                    }
                    if ((Test.Data("SSN") == null) || (Test.Data("SSN") == ""))
                    {
                        Logger.info("SSN Field is empty");
                    }
                    else
                    {
                        Browser.Click("WhatToDoHerePage", "SearchByDropDown");
                        Browser.Click("WhatToDoHerePage", "SearchBySSN");
                        Browser.EnterText("WhatToDoHerePage", "KeywordTextBox", Test.Data("SSN"));
                    }
                    Browser.Click("WhatToDoHerePage", "SearchButton");
                    Browser.WaitforthePageLoad();
                    Browser.IsDisplayed("EmployeeSearchResultsPage", "RadioButton");
                    Browser.VerifyText("EmployeeSearchResultsPage", "SelectLabel", "SELECT");
                    Browser.VerifyText("EmployeeSearchResultsPage", "SSNLabel", "SSN");
                    Browser.VerifyText("EmployeeSearchResultsPage", "LastNameLabel", "LAST NAME");
                    Browser.VerifyText("EmployeeSearchResultsPage", "FirstNameLabel", "FIRST NAME");
                    Browser.VerifyText("EmployeeSearchResultsPage", "StatusLabel", "STATUS");
                    Browser.VerifyText("EmployeeSearchResultsPage", "CityLabel", "CITY");
                    Browser.VerifyText("EmployeeSearchResultsPage", "StateLabel", "STATE");
                    Browser.VerifyText("EmployeeSearchResultsPage", "LOBLabel", "LOB");
                    Browser.IsDisplayed("EmployeeSearchResultsPage", "ImpersonateButton");
                    Browser.Click("EmployeeSearchResultsPage", "RadioButton");
                    Browser.Click("EmployeeSearchResultsPage", "ImpersonateButton");
                    Browser.WaitforthePageLoad();

                    Browser.VerifyText("ImpersonatedHomePage", "AttensionLabel", "ATTENTION");
                    Browser.VerifyText("ImpersonatedHomePage", "ImpersonatingLabel", "You are currently impersonating");
                    Browser.IsDisplayed("ImpersonatedHomePage", "EndImpersonationButton");

                }
                else
                {
                    Test.TakeScreenShot();
                    Test.Result("URL should be present in Test Data Sheet", "URL is not present in Test Data Sheet Please Update", "Failed");
                }

            }
            catch (Exception e)
            {
                Logger.error("Impersonate Function has the following exception" + e);
                Test.Result("Impersonation function should be performed", "Impersonation function is not performed", "Failed");
                Test.TakeScreenShot();
            }
            finally
            {
                Logger.info("Exiting from 'ImpersonateFunc' method");
            }

        }
        public void EndImpersonationFunc()
        {
            Logger.info("Entering to 'EndImpersonationFunc' method");
            try
            {
                Browser.Click("ImpersonatedHomePage", "EndImpersonationButton");
                Browser.VerifyPageURLContains("ImpersonatedHomePage", "error");
                Browser.WaitforthePageLoad();
                Browser.IsDisplayed("loginPage", "LoginButton");
            }
            catch (Exception e)
            {
                Logger.error("Impersonate Function has the following exception" + e);
                Test.Result("Impersonation function should be performed", "Impersonation function is not performed", "Failed");
                Test.TakeScreenShot();
            }
            finally
            {
                Logger.info("Exiting from 'EndImpersonationFunc' method");
            }
        }
    }
}