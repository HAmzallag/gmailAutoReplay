using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Net.Mail;

namespace GmailAutoReplay
{
    class Gmail
    {
        static string[] Scopes = { GmailService.Scope.GmailSend, GmailService.Scope.GmailModify, GmailService.Scope.GmailReadonly };
        static string ApplicationName = "Gmail automatic responser";
        public static string GoogleAccessToken { get; private set; }
        private UserCredential credential;
        private GmailService service;

        public Gmail(string clientJSON)
        {
            using (var stream = new FileStream(clientJSON, FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/gmail-dotnet-quickstart");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }

            service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        public Message getLastMessage(DateTime dateDebut)
        {
            try
            {
                UsersResource.MessagesResource.ListRequest req = service.Users.Messages.List("me");
                req.LabelIds = "UNREAD";
                IList<Message> mes;
                if ((mes = req.Execute().Messages) == null)
                    return null;
                Message retour = service.Users.Messages.Get("me", mes.First().Id).Execute();

                //verifie la date quand celle ci est présente          
                long dateDebutTimeStamp = (long)(dateDebut.Date.ToUniversalTime() - new DateTime(1970, 1, 1)).TotalMilliseconds;
                if (retour.InternalDate != null && retour.InternalDate < dateDebutTimeStamp)
                    return null;

                ModifyMessageRequest mods = new ModifyMessageRequest();
                var label = new List<string>();
                label.Add("UNREAD");
                mods.RemoveLabelIds = label;
                service.Users.Messages.Modify(mods, "me", retour.Id).Execute();
                return retour;
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
                return null;
            }
        }

        public string sendMessage(string emailAdress, string destEmail, string subject, string body, List<string> attachments)
        {
            try
            {
                var msg = new AE.Net.Mail.MailMessage
                {
                    Subject = subject,
                    Body = body,
                    From = new MailAddress(emailAdress)
                };

                foreach (string path in attachments)
                {
                    byte[] bytes = File.ReadAllBytes(path);
                    AE.Net.Mail.Attachment file = new AE.Net.Mail.Attachment(bytes, GetMimeType(path), Path.GetFileName(path), true);
                    msg.Attachments.Add(file);
                }

                msg.To.Add(new MailAddress(destEmail));
                msg.ReplyTo.Add(msg.From); // Bounces without this!!
                var msgStr = new StringWriter();
                msg.Save(msgStr);

                var result = service.Users.Messages.Send(new Message
                {
                    Raw = Base64UrlEncode(msgStr.ToString())
                }, "me").Execute();
                Console.WriteLine("{0} : Message sent to {1}.", DateTime.Now, destEmail);
                return (DateTime.Now + " : Message sent to : " + destEmail + ".");
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
                return ("An error occurred: " + e.Message);
            }
        }

        private static string GetMimeType(string fileName)
        {
            string mimeType = "application/unknown";
            string ext = Path.GetExtension(fileName).ToLower();
            Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
            if (regKey != null && regKey.GetValue("Content Type") != null)
                mimeType = regKey.GetValue("Content Type").ToString();
            return mimeType;
        }

        private string Base64UrlEncode(string input)
        {
            var inputBytes = Encoding.UTF8.GetBytes(input);
            // Special "url-safe" base64 encode.
            return Convert.ToBase64String(inputBytes)
              .Replace('+', '-')
              .Replace('/', '_')
              .Replace("=", "");
        }
    }
}