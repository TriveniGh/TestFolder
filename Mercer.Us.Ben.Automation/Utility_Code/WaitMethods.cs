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
    public class WaitMethods:Test
    {

        public static bool WaitForElementNotPresent()
        {
            return true;
        }

        public static bool WaitForElementVisible(By element, int timeInSeconds = Constants.ImplicitTimeout)
        {
            Logger.info("Entering to 'WaitForElementVisible' method" + "Waiting for element to become visible: " + element.ToString());
            TimeSpan timer = new TimeSpan(0, 0, timeInSeconds);
            WebDriverWait wait = new WebDriverWait(driver, timer);
            try
            {
                wait.Until(ExpectedConditions.ElementIsVisible(element));
                return true;
            }
            catch (Exception e)
            {
                Logger.error("Method name 'WaitForElementVisible' throws  exception", e);
                return false;
            }
            finally
            {
                Logger.info("Exiting from 'WaitForElementVisible' method");
            }
            
        }







    }
}
