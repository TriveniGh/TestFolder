using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mercer.Us.Ben.Automation.Main;
using OpenQA.Selenium;
using Mercer.Us.Ben.Automation.Utility_Code;

namespace Mercer.Us.Ben.Automation.Test_Methods.MBC
{
    public class MBC_Login : Test
    {

        public void LoginPageValidationfunc()
        {
            Logger.info("Entering to 'LoginPageValidationfunc' method");
            try
            {
                if (Test.Data("URL") != "" || Test.Data("URL") != null)
                {
                    if (Browser.Launch(Test.Data("URL")))
                    {
                        Browser.WaitforthePageLoad();
                        if (Test.Data("SSO") == "N")
                        {
                            Browser.IsDisplayed("MBC_Login", "LoginImage");
                            Browser.VerifyText("MBC_Login", "ReturningUserLabel", "Returning Users");
                            Browser.VerifyText("MBC_Login", "LogInLabel", "Log in to your existing account.");
                            Browser.IsDisplayed("MBC_Login", "MBCUserName");
                            Browser.IsDisplayed("MBC_Login", "MBCPassword");
                            Browser.IsDisplayed("MBC_Login", "LoginButton");
                            Browser.VerifyText("MBC_Login", "NewUsersLabel", "New Users");
                            Browser.VerifyText("MBC_Login", "RegisterUrAccLabel", "Register your account now.");
                            Browser.IsDisplayed("MBC_Login", "GetStartedButton");
                            Browser.IsDisplayed("MBC_Login", "HelpfulHintsLink");
                            Browser.IsDisplayed("MBC_Login", "RecommendedBrowLink");
                            Browser.IsDisplayed("MBC_Login", "PrivacyPolicyLink");
                            Browser.IsDisplayed("MBC_Login", "TermsofUseLink");
                            //  Test.TakeScreenShot();
                        }
                        else
                        {
                            //  Test.TakeScreenShot();
                            Test.Result("Login Page should not be validated because home page directly opens", "Homepage opened directly", "Passed");
                        }

                    }
                }
                else
                {
                  //  Test.TakeScreenShot();
                    Test.Result("URL should be present in Test Data Sheet", "URL is not present in Test Data Sheet Please Update", "Failed");
                }
            }
            catch (NoSuchElementException e)
            {
                Logger.error("Method name: 'LoginPageValidationfunc' has NoSuchElementException", e);
               // Test.TakeScreenShot();
                Test.Result("LoginPage should be validated","LoginPage is not valid","Failed");
            }
            finally
            {
                Logger.info("Exiting from 'LoginPageValidationfunc' method");
            }
        }

        public void PrivacyPolicyMainfunc()
        {
            Logger.info(" Entering to 'PrivacyPolicyMainfunc' method execution");
            try
            {
                Browser.Click("MBC_Login", "PrivacyPolicyLink");
                Browser.SwitchToNewWindow("PrivacyPolicyPage");
                Browser.WaitforthePageLoad();
                Browser.VerifyText("PrivacyPolicyPage", "PrivacyPolicyLabel", "Privacy Policy");
                //Browser.VerifyText("PrivacyPolicyPage", "PrivacyContentLabel", "This is Privacy Policy Text.");
                Browser.WaitforthePageLoad();
        //        Test.TakeScreenShot();
                Browser.SwitchToLastWindow("MBC_Login");
            }
            catch (Exception e)
            {
                Logger.error("Privacy Policy Link function contains the following exception" + e);
                Test.Result("Privacy Policy Link Validation should be performed", "Privacy Policy Link is not Valid", "Failed");
             //   Test.TakeScreenShot();
            }
            finally
            {
                Logger.info(" Exiting from 'PrivacyPolicyMainfunc' method");
            }
        }

        public void TermsOfUseMainfunc()
        {
            Logger.info(" Entering to 'TermsOfUseMainfunc' method execution");
            try
            {
                Browser.Click("MBC_Login", "TermsofUseLink");
                Browser.SwitchToNewWindow("TermsOfUsePage");
                Browser.WaitforthePageLoad();
                Browser.VerifyText("TermsOfUsePage", "TermsOfUseLabel", "Terms of Use");
             //   Test.TakeScreenShot();
                Browser.SwitchToLastWindow("MBC_Login");
                //Browser.VerifyText("TermsOfUsePage", "TermsOfUseContent", "Terms of Use");
            }
            catch (Exception e)
            {
              //  Test.TakeScreenShot();
                Logger.error("Terms Of Use Link function contains the following exception", e);
                Test.Result("Validation of Terms of use Link should be performed ", "Terms of use Link is not valid", "Failed");
            }
            finally
            {
                Logger.info(" Exiting from 'TermsOfUseMainfunc' method");
            }
        }

        public void HelpfulHintsMainfunc()
        {
            Logger.info(" Entering to 'HelpfulHintsMainfunc' method execution");
            try
            {
                Browser.Click("MBC_Login", "HelpfulHintsLink");
                Browser.SwitchToNewWindow("HelpfulHintsPage");
                Browser.WaitforthePageLoad();
                Browser.VerifyText("HelpfulHintsPage", "HelpfulLinkLabel", "Helpful hints for accessing your account");
                Browser.SwitchToLastWindow("MBC_Login");
            }
            catch (Exception e)
            {
                Logger.error("Helpful Hints Link function contains the following exception" + e);
                Test.Result("Helpful Hints Page Validation should be performed", "Helpful Hints Page is not valid", "Failed");
            //    Test.TakeScreenShot();
            }
            finally
            {
                Logger.info(" Exiting from 'HelpfulHintsMainfunc' method");
            }
        }

        public void RecommendedBrowserMainfunc()
        {
            Logger.info(" Entering to 'RecommendedBrowserMainfunc' method execution");
            try
            {
                Browser.Click("MBC_Login", "RecommendedBrowLink");
                Browser.SwitchToNewWindow("RecommendedBrowserPage");
                Browser.WaitforthePageLoad();
                Browser.IsDisplayed("RecommendedBrowserPage", "SiteLink");
                Browser.VerifyText("RecommendedBrowserPage", "RecommendedBrowserLabel", "Recommended browsers");
                Browser.VerifyText("RecommendedBrowserPage", "MicrosoftIELabel", "Microsoft Internet Explorer 8, 9, or 10");
                Browser.VerifyText("RecommendedBrowserPage", "FirefoxLabel", "Firefox*");
                Browser.VerifyText("RecommendedBrowserPage", "SafariLabel", "Safari 5, or 6");
                Browser.VerifyText("RecommendedBrowserPage", "ChromeLbael", "Chrome 30");
          //      Test.TakeScreenShot();
                Browser.SwitchToLastWindow("MBC_Login");

            }
            catch (Exception e)
            {
                Logger.error("Recommended Browsers Link function contains the following exception" + e);
          //      Test.TakeScreenShot();
                Test.Result("Recommended Browser Link Validation should be performed", "Recommended Browser Link is not valid", "Failed");
            }
            finally
            {
                Logger.info(" Exiting from 'RecommendedBrowserMainfunc' method");
            }
        }

        public void ReturningUserfunc()
        {
            Logger.info(" Entering to 'ReturningUserfunc' method execution");
            try
            {
                if (Test.Data("SSO") == "N")
                {
                   Browser.WaitforthePageLoad();
                   Browser.EnterText("MBC_Login","MBCUserName",Test.Data("UserName"));
                   Browser.EnterText("MBC_Login","MBCPassword",Test.Data("Password"));
                   Browser.IsDisplayed("MBC_Login","LoginButton");
                   Browser.Click("MBC_Login","LoginButton");
                   Browser.WaitforthePageLoad();
                   Browser.IsDisplayed("HomePage","EstimateNowButton");
                   Browser.WaitforthePageLoad();
                  
                }
                else
                {
                   Browser.WaitforthePageLoad();
                   Browser.IsDisplayed("HomePage","EstimateNowButton");
                }
            }
            catch (Exception e)
            {
                Logger.error("ReturningUser function throws this exception : ", e);
            //    Test.TakeScreenShot();
                Test.Result("Page should be loaded ", "Waiting for page load", "Failed");
            }
            finally
            {
                Logger.info(" Exiting from 'ReturningUserfunc' method");
            }
        }
        
        public void ErrorPageMainfunc()
        {
            try
            {
                if (Browser.IsDisplayed("ErrorMsg", "ContinueButton"))
                {
                    Browser.Click("ErrorMsg", "ContinueButton");
                    Browser.Click("HomePage", "MenuButton");
                    Browser.Click("HomePage", "HomeLink");
                }

             //   Test.TakeScreenShot();
            }
            catch (Exception e)
            {
                Logger.error("Continue Button function contains the following exception" + e);
            }
        }

        public void LogoutFunc()
        {
            try
            {
                Browser.IsDisplayed("HomePage", "LogoutLink");
                Browser.JavsScriptClick("HomePage", "LogoutLink");
                if (driver.Url.ToLower().Contains("error"))
                {
                    Test.Result("Logout should be successful", "Un expected error while performing logout", "Failed");
                    Test.TakeScreenShot();
                }
                //Browser.WaitforthePageLoad();
                Browser.IsDisplayed("MBC_Login", "LoginButton");
                Browser.WaitforthePageLoad();
           //     Test.TakeScreenShot();
            }
            catch (Exception e)
            {
                Logger.error("Logout Function has the following exception" + e);
                Test.Result("Logout Link Should be clicked in home page", "Logout Link is Clicked but not navigated to Login Page", "Failed");
              //  Test.TakeScreenShot();
            }
        }

        


    }
}

