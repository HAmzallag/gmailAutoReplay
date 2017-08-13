using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Gmail.v1.Data;

namespace GmailAutoReplay
{
    public class Autoreplay
    {
        private const string PASS = "passpartout";
        private string clientJson = "";
        private string emailBox = "";
        private string subject = "";
        private string body = "";
        private List<string> attachments;

        public Autoreplay(string ClientJsonAuth, string EmailBox, string Subject, string Body, List<string> Attachments, string pass, DateTime dateDebut, string DemoEmail)
        {
            if (pass != PASS)
                return;
            clientJson = ClientJsonAuth;
            emailBox = EmailBox;
            subject = Subject;
            body = Body;
            attachments = Attachments;
            Gmail gmail = new Gmail(clientJson);
            Message lastUnreadMessage;
            if ((lastUnreadMessage = gmail.getLastMessage(dateDebut)) == null)
                return;

            string email = lastUnreadMessage.Payload.Headers.Where(h => h.Name == "From").First().Value;
            int pFrom = email.IndexOf("<") + "<".Length;
            int pTo = email.LastIndexOf(">");
            email = email.Substring(pFrom, pTo - pFrom);

            if (DemoEmail.Length >= 3)
            {
                gmail.sendMessage(emailBox, DemoEmail, subject, body, attachments);
            }
            else
            {
                gmail.sendMessage(emailBox, email, subject, body, attachments);
            }

        }

        public Autoreplay(string ClientJsonAuth, string EmailBox, string Subject, string Body, List<string> Attachments, string pass, DateTime dateDebut)
    : this(ClientJsonAuth, EmailBox, Subject, Body, Attachments, pass, dateDebut, "")
        {
        }

        public Autoreplay(string ClientJsonAuth, string EmailBox, string Subject, string Body, List<string> Attachments, string pass)
    : this(ClientJsonAuth, EmailBox, Subject, Body, Attachments, pass, DateTime.Now.AddDays(-1), "")
        {
        }

        public Autoreplay(string ClientJsonAuth, string EmailBox, string Subject, string Body, List<string> Attachments, string pass, string demoEmail)
   : this(ClientJsonAuth, EmailBox, Subject, Body, Attachments, pass, DateTime.Now.AddDays(-1), demoEmail)
        {
        }

    }
}
