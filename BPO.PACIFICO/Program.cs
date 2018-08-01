using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GmailQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/gmail-dotnet-quickstart.json
        static string[] Scopes = { GmailService.Scope.GmailReadonly };
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

            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List("bponaa@gmail.com");
            request.MaxResults = 5;
            request.LabelIds = "INBOX";
            request.IncludeSpamTrash = false;
            request.Q = "is:unread";
            
            // List Messages.
            IList<Message> messages = request.Execute().Messages;
            Console.WriteLine("Messages:");
            if (messages != null && messages.Count > 0)
            {
                foreach (var message in messages)
                {
                    var infoRequest = service.Users.Messages.Get("bponaa@gmail.com", message.Id);
                    var infoResponse = infoRequest.Execute();
                    if (infoResponse != null)
                    {
                        String from = "";
                        String date = "";
                        String subject = "";
                        String body = "";

                        foreach (var mParts in infoResponse.Payload.Headers)
                        {
                            if (mParts.Name == "Date")
                            {
                                date = mParts.Value;
                            }
                            else if (mParts.Name == "From")
                            {
                                from = mParts.Value;
                            }
                            else if (mParts.Name == "Subject")
                            {
                                subject = mParts.Value;
                            }

                            if (date != "" && from != "")
                            {
                                if (infoResponse.Payload.Parts == null && infoResponse.Payload.Body != null)
                                {
                                    body = infoResponse.Payload.Body.Data;
                                }
                                else
                                {
                                    body = getNestedParts(infoResponse.Payload.Parts, "");
                                }
                                //need to replace some characters as the data for the email's body is base64
                                String codedBody = body.Replace("-", "+");
                                codedBody = codedBody.Replace("_", "/");
                                byte[] data = Convert.FromBase64String(codedBody);
                                body = Encoding.UTF8.GetString(data);


                                //now you have the data you want....

                            }

                        }
                    }

                    var abc = message.Payload.Headers;
                    Console.WriteLine("{0}", abc);
                }
            }
            else
            {
                Console.WriteLine("No messages found.");
            }

            /*
            // Define parameters of request.
            UsersResource.LabelsResource.ListRequest request = service.Users.Labels.List("me");

            // List labels.
            IList<Label> labels = request.Execute().Labels;
            Console.WriteLine("Labels:");
            if (labels != null && labels.Count > 0)
            {
                foreach (var labelItem in labels)
                {
                    Console.WriteLine("{0}", labelItem.Name);
                }
            }
            else
            {
                Console.WriteLine("No labels found.");
            }
            */
            Console.Read();
        }

        static String getNestedParts(IList<MessagePart> part, string curr)
        {
            string str = curr;
            if (part == null)
            {
                return str;
            }
            else
            {
                foreach (var parts in part)
                {
                    if (parts.Parts == null)
                    {
                        if (parts.Body != null && parts.Body.Data != null)
                        {
                            str += parts.Body.Data;
                        }
                    }
                    else
                    {
                        return getNestedParts(parts.Parts, str);
                    }
                }

                return str;
            }

        }
    }
}