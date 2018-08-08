using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BPO.PACIFICO.NOTIFICACION.EMAIL
{
    class Program
    {
        //static string[] Scopes = { GmailService.Scope.GmailSend, GmailService.Scope.GmailReadonly};
        static string[] Scopes = { GmailService.Scope.GmailSend };
        static string ApplicationName = "Gmail API .NET Quickstart";

        static void Main(string[] args)
        {


            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            //// Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


            string emails = "bponaa@gmail.com,luistrujilloh@hotmail.com";


            string plainText = "To:" + emails + " \r\n" +
                                "Subject: subject Test\r\n" +
                                "Content-Type: text/html; charset=us-ascii\r\n\r\n" +
                                "<h1>Body Tekykjhkst </h1>";

            var newMsg = new Google.Apis.Gmail.v1.Data.Message();
            newMsg.Raw = Program.Base64UrlEncode(plainText.ToString());
            service.Users.Messages.Send(newMsg, "bponaa@gmail.com").Execute();
        }

        public static string Base64UrlEncode(string input)
        {
            var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(inputBytes).Replace("+", "-").Replace("/", "_").Replace("=", "");
        }


    }
}
