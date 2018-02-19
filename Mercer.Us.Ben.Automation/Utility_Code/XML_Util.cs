using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Mercer.Us.Ben.Automation.Main;
using System.IO;

namespace Mercer.Us.Ben.Automation.Utility_Code
{
    public class XML_Util : Test
    {
        public static By readObjectData(string pageName, string elementName)
        {
            XmlDocument xd = new XmlDocument();
            xd.PreserveWhitespace = true;

            string Xml_Path = Test.CurrentProjectPath + "Mercer.Us.Ben.Automation" + Path.DirectorySeparatorChar + "Repository" + Path.DirectorySeparatorChar + Test.Sheet_Name + ".xml";
            xd.Load(Xml_Path);

            // string ObjectName = "tour ";
            // string Module = "flickPage";
            XmlNodeList nodelist = xd.SelectNodes("/pages/page[@name='" + pageName + "']/uiobject[@name='" + elementName + "']/locator"); // get all <testcase> nodes
            //
            string returnValue = "";
            foreach (XmlNode node in nodelist) // for each <testcase> node
            {

                returnValue = node.ChildNodes.Item(0).InnerText;
                Console.WriteLine(returnValue);

            } // foreach <testcase> node

            By webElementLocatorBy = findBy(returnValue);
        
            
            return webElementLocatorBy;
        }

        public static IWebElement findTheElement(string webElementLocator)
        {
            By webElementLocatorBy = findBy(webElementLocator);
            try
            {
                return driver.FindElement(webElementLocatorBy);
            }
            catch (Exception e)
            {
                Logger.error("Unable to find element with locator " + webElementLocator + ". Trying to wait for element to load." + e);

                // Wait for the element to load before doing a find element
                //   waitForPageElementToLoad(webElementLocator, TestConstants.MEDIUM_WAIT, WAIT_SLEEP);
                return driver.FindElement(webElementLocatorBy);
            }
        }


        public static By findBy(string input)
        {
            string returnVal = input;
            By locatorObj = null;

            if (input != null)
            {
                int cdataIndex = input.IndexOf("CDATA[");
                if (cdataIndex > 0)
                {
                    returnVal = input.Substring(cdataIndex + 1);
                    int cdataStart = returnVal.IndexOf("[");
                    int cdataEnd = returnVal.IndexOf("]");

                    if (cdataStart > 0 && cdataEnd > 0)
                    {
                        returnVal = (returnVal.Substring(cdataStart + 1, cdataEnd)).Trim();
                    }
                    int index = returnVal.IndexOf("=");

                    if (index > 0)
                    {
                        returnVal = returnVal.Substring(index + 1);
                    }

                }
                else
                {
                    int index = input.IndexOf("=");
                    if (index > 0)
                    {
                        returnVal = input.Substring(index + 1);
                    }
                }

                if (returnVal != null)
                    returnVal = returnVal.Trim();
                if (input.StartsWith("name="))
                {
                    locatorObj = By.Name(returnVal);
                }
                else if (input.StartsWith("id="))
                {
                    locatorObj = By.Id(returnVal);
                }
                else if (input.StartsWith("xpath="))
                {
                    locatorObj = By.XPath(returnVal);
                }
                else if (input.StartsWith("css="))
                {
                    locatorObj = By.CssSelector(returnVal);
                }
                else if (input.StartsWith("class="))
                {
                    locatorObj = By.ClassName(returnVal);
                }
                else if (input.StartsWith("link="))
                {
                    locatorObj = By.LinkText(returnVal);
                }
                else if (input.StartsWith("partialLink="))
                {
                    locatorObj = By.PartialLinkText(returnVal);
                }
                else if (input.StartsWith("tag="))
                {
                    locatorObj = By.TagName(returnVal);
                }
            }

            return locatorObj;
        }
    }
}
