using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using TechTalk.SpecFlow;
using TestAutomation.Framework.Helpers.Assertions;
using TestAutomation.Framework.Helpers.Files;
using TestAutomation.Framework.Helpers.Loggers;
using TestAutomation.Framework.Helpers.Outlook;
using TestAutomation.Framework.Helpers.Strings;
using TestAutomation.Framework.Helpers.UI_Helpers;
using TestAutomation.OutlookTests.Pages.PagesData;

namespace TestAutomation.OutlookTests.StepDefiinitions
{
    [Binding]
    public sealed class OutlookSteps : UIFramework
    {
        static OutlookHelper outlookHelper = new OutlookHelper();
        static Microsoft.Office.Interop.Outlook.MAPIFolder inbox = outlookHelper.GetOutlookFolder("Inbox");
        readonly FileExtentions fileExtentions = new FileExtentions();
        OutlookPageData outlookPageData;
        public OutlookSteps()
        {
            //UIController.Instance.Driver = driver;
            outlookPageData = Page<OutlookPageData>();
        }

        [Given(@"I send an email to (.*)")]
        public void GivenISendAnEmail(string emailId)
        {
            ConfigureEmail();
        }

        [Then(@"I recieve an email in my inbox")]
        public void ThenIRecieveAnEmailInMyInbox()
        {
            // Verify with filter property Subject and Body 
            outlookHelper.VerifyOutlookEmails("[Subject] = 'Test Automation'", "Inbox", string.Empty, string.Empty, "This is an test email.");
        }

        [Then(@"I can get email items sorted by recieved date")]
        public void ThenICanGetEmailItemsSortedByRecievedDate()
        {
            Items items = outlookHelper.SortOutlookEmails("ReceivedTime", "[Subject] = 'Test Automation'", "Inbox");

            for (int i = 1; i <= items.Count; i++)
            {
                var item = items[i];
                AssertHelpers.AssertContains(item.Subject.Replace(System.Environment.NewLine, string.Empty), "Test Automation");
                AssertHelpers.AssertContains(item.Body.Replace(System.Environment.NewLine, string.Empty), "This is an test email.");
            }
        }

        [Then(@"I can verify email has attahcments")]
        public void ThenICanVerifyEmailHasAttahcments()
        {
            outlookHelper.VerifyHasAttachments("TestFile.txt", "[Subject] = 'Test Automation'");
        }


        [Then(@"I can get attachment from mail")]
        public void ThenICanGetAttachmentFromMail()
        {
            var attachmentList = outlookHelper.GetAttachments("[Subject] = 'Test Automation'");
        }

        [Then(@"I can save attachment from mail")]
        public void ThenICanSaveAttachmentFromMail()
        {
            outlookHelper.SaveAttachments("[Subject] = 'Test Automation'");
        }

        [Then(@"I can read the recived email")]
        public void ThenICanReadTheRecivedEmail()
        {
            outlookHelper.ReadOutlookEmail(inbox, "[Subject] = 'Test Automation'");

            var from = outlookHelper.From;
            var to = outlookHelper.To;
            var subject = outlookHelper.Subject;

            var body = outlookHelper.Body;

            AssertHelpers.AssertEquals(from, "Captain.America@Hotmail.com");
            AssertHelpers.AssertEquals(to, "Captain.America@Hotmail.com");
            AssertHelpers.AssertEquals(subject, "Test Automation");
            AssertHelpers.AssertContains(body, "This is an test email.");
        }

        [Given(@"I have saved email")]
        public void GivenIHaveSavedEmail()
        {
            string folderPath = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)) +
                           ConfigurationManager.AppSettings["Files"];

            fileExtentions.VerifyFileExists(folderPath, "Test Automation");
        }

        [Then(@"I can read saved email")]
        public void ThenICanReadSavedEmail()
        {
            outlookHelper.ReadSavedEmail("Test Automation.msg");

            var from = outlookHelper.From;
            var to = outlookHelper.To;
            var subject = outlookHelper.Subject;

            var body = outlookHelper.Body;

            AssertHelpers.AssertEquals(from, "Captain.America@Hotmail.com");
            AssertHelpers.AssertEquals(to, "Captain.America@Hotmail.com");
            AssertHelpers.AssertEquals(subject, "Test Automation");
            AssertHelpers.AssertContains(body, "This is an test email.");
        }

        [Then(@"I can read the recived email body line by line")]
        public void ThenICanReadTheRecivedEmailBodyLineByLine()
        {
            Items items = outlookHelper.SortOutlookEmails("ReceivedTime", "[Subject] = 'Test Automation'");

            string line = string.Empty;
            string githublink = string.Empty;

            var sr = new StringReader(((MailItem)items[items.Count]).Body);
            while ((line = sr.ReadLine()) != null)
            {
                if (line.Contains("Github"))
                {
                    githublink = line.Substring(line.LastIndexOf("Github"),6);
                    break;
                }
            }
        }

        [Then(@"I can click on link in the recived email")]
        public void ThenICanClickOnLinkInTheRecivedEmail()
        {
            IWebDriver driver = new ChromeDriver();   
            UIController.Instance.Driver = driver;
            outlookHelper.ReadOutlookEmail(inbox, "[Subject] = 'Test Automation'");

            var htmlLinks = outlookHelper.HTMLLinkList;
            var match = htmlLinks.FirstOrDefault(stringToCheck => stringToCheck.Contains("github.com"));
            UIActions.NavigateToUrl(match.ToString());
        }

        public void ConfigureEmail()
        {
            // Command-line argument must be the SMTP host.
            string _sender = outlookPageData.SenderEmail;// "Captain.America@hotmail.com";
            string _password = outlookPageData.SenderPassword;

            SmtpClient client = new SmtpClient(outlookPageData.SmptClient);

            client.Port = 587;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            System.Net.NetworkCredential credentials =
                new System.Net.NetworkCredential(_sender, _password);
            client.EnableSsl = true;
            client.Credentials = credentials;

            MailMessage message = new MailMessage(_sender, outlookPageData.RecieverEmail);
            message.Subject = outlookPageData.Subject;
            message.IsBodyHtml = true;
            message.Body = outlookPageData.Body;
            message.CC.Add(outlookPageData.Cc);
            message.Attachments.Add(new System.Net.Mail.Attachment(outlookPageData.Attachments));
            client.Send(message);
            Report.AddInfo("Email with subject " + outlookPageData.Subject + " has been sent to reciepient: " + outlookPageData.RecieverEmail);
        }
    }
}
