using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestAutomation.Framework.Helpers.Json;
using TestAutomation.Framework.Helpers.Setup;

namespace TestAutomation.OutlookTests.Pages.PagesData
{
    public class OutlookPageData : BasePage<OutlookPageData>
    {
        readonly JsonHelper jHelper = new JsonHelper();

        public string CurrentFileDirectory;

        private const string smptClient = "SmptClient";
        public string SmptClient => jHelper.GetLocalData(smptClient, "CommonData");

        private const string senderEmail = "SenderEmail";
        public string SenderEmail => jHelper.GetLocalData(senderEmail, "CommonData");

        private const string senderPassword = "SenderPassword";
        public string SenderPassword => jHelper.GetLocalData(senderPassword, "CommonData");

        private const string recieverEmail = "RecieverEmail";
        public string RecieverEmail => jHelper.GetLocalData(recieverEmail, "CommonData");

        private const string cc = "CcEmail";
        public string Cc => jHelper.GetLocalData(cc, "CommonData");

        private const string subject = "Subject";
        public string Subject => jHelper.GetLocalData(subject, "CommonData");

        private const string body = "Body";
        public string Body => jHelper.GetLocalData(body, "CommonData");

        private const string attachments = "Attachment";

        public string Attachments
        {
            get
            {
                CurrentFileDirectory = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)));
                CurrentFileDirectory = CurrentFileDirectory + jHelper.GetLocalData(attachments, "CommonData");
                return CurrentFileDirectory.Replace("file:\\", "");
            }
        }
    }
}
