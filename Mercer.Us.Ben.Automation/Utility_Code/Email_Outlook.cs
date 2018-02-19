using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using Mercer.Us.Ben.Automation.Main;
namespace Mercer.Us.Ben.Automation.Utility_Code
{
   public  class Email_Outlook :Test
    {

        public static void sendEMailThroughOUTLOOK()
        {
            try
            {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                // Create a new mail item.
                Outlook._MailItem oMsg = (Outlook._MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody. 
                //add the body of the email
                //      "Dear Admin, You have recieved an enquiry onyour Website. Following are the details of enquiry:<br/><br/>Name: " + TextBox2.Text + "<br/>Address: " + TextBox3.Text + ", " + TextBox4.Text + "<br/>Phone: " + TextBox5.Text + "<br/>Email: " + TextBox2.Text + "<br/>Query: " + TextBox6.Text+"<br/> Regards, <br/> Veritas Team"
                oMsg.HTMLBody = "Hi,\n \n <br/> <br/>Automation Execution has been started \n<br/><br/> Thanks & Regards \n<br/> Raghu Ram Reddy<br/><br/><br/>***This is an Auto generated mail***";
                //  oMsg.HTMLBody = "Automation Execution has been started!!";
                //Add an attachment.
                String sDisplayName = "MyAttachment";
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //   now attached the file
                Outlook.Attachment oAttach = oMsg.Attachments.Add(@"D:\\Suite_Driver.xlsx", iAttachType, iPosition, sDisplayName);
                //Subject line

                oMsg.Subject = "Automation Execution has been started!!";

                //    Outlook.MailItem mail;


                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.

                //    Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("raghuram.reddy@lntinfotech.com");
                //    oRecip.Resolve();

                Outlook.Recipient oRecip2 = (Outlook.Recipient)oRecips.Add("Dhamayanthi.Amirthalingam@lntinfotech.com");
                oRecip2.Resolve();

                //Outlook.Recipient oRecip3 = (Outlook.Recipient)oRecips.Add("Abishek.AK@lntinfotech.com");
                //oRecip3.Resolve();

                // Send.
                oMsg.Send();
                // Clean up.
                oRecip2 = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
            }//end of try block
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }//end of catch
        }//end of Email Method


        public static void send_Report_EMailThroughOUTLOOK()
        {
            try
            {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                // Create a new mail item.
                Outlook._MailItem oMsg = (Outlook._MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody. 
                //add the body of the email
                //      "Dear Admin, You have recieved an enquiry onyour Website. Following are the details of enquiry:<br/><br/>Name: " + TextBox2.Text + "<br/>Address: " + TextBox3.Text + ", " + TextBox4.Text + "<br/>Phone: " + TextBox5.Text + "<br/>Email: " + TextBox2.Text + "<br/>Query: " + TextBox6.Text+"<br/> Regards, <br/> Veritas Team"
                oMsg.HTMLBody = "Hi,\n \n <br/> <br/>Automation Execution has been completed. \n<br/>Please Find the attached Automation result report to this mail.\n \n <br/> <br/> Thanks & Regards \n<br/> Automation Team<br/><br/><br/>***This is an Auto generated mail ***";
                //  oMsg.HTMLBody = "Automation Execution has been started!!";
                //Add an attachment.
                String sDisplayName = "MyAttachment";
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //   now attached the file
                Outlook.Attachment oAttach = oMsg.Attachments.Add(Test.OutputfilePath, iAttachType, iPosition, sDisplayName);
                //Subject line

                oMsg.Subject = "Automation Execution Report!!";
                //    Outlook.MailItem mail;
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.
                //  Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("raghuram.reddy@lntinfotech.com");
                //  oRecip.Resolve();

                Outlook.Recipient oRecip2 = (Outlook.Recipient)oRecips.Add("Dhamayanthi.Amirthalingam@lntinfotech.com");
                oRecip2.Resolve();
                // Outlook.Recipient oRecip2 = (Outlook.Recipient)oRecips.Add("Dhamayanthi.Amirthalingam@lntinfotech.com");
                // oRecip2.Resolve();
                //Outlook.Recipient oRecip3 = (Outlook.Recipient)oRecips.Add("Abishek.AK@lntinfotech.com");
                //oRecip3.Resolve();

                // Send.
                oMsg.Send();
                // Clean up.
                oRecip2 = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
            }//end of try block
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }//end of catch
        }//end of Email Method

    }
}
