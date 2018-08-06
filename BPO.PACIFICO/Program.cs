using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
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
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using everis.Ees.Proxy.Services;
using BPO.PACIFICO;

namespace GmailQuickstart
{
    class Program : IRobot
    {
        static BaseRobot<Program> _robot = null;
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/gmail-dotnet-quickstart.json
        static string[] Scopes = { GmailService.Scope.GmailReadonly };
        static string[] _palabras = new string[2];
        static string ApplicationName = "Gmail API .NET Quickstart";
        static List<Puntuacion> puntos = new List<Puntuacion>();
        static void Main(string[] args)
        {
            _robot = new BaseRobot<Program>(args);
            _robot.Start();



        }

        protected override void Start()
        {



            Email();
        }

        public void Email()
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
                    Message infoResponse = service.Users.Messages.Get("bponaa@gmail.com", message.Id).Execute();
                    AcionRequest(infoResponse);

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
        public void AcionRequest(Message infoResponse)
        {

            if (infoResponse != null)
            {
                String from = String.Empty;
                String date = String.Empty;
                String subject = String.Empty;
                String body = String.Empty;

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



                        //now you have the data you want....
                        EvaluarPuntuacion(String.Concat(decodeBase64(body), " ", subject));



                        date = String.Empty;
                        from = String.Empty;
                        break;
                    }

                }
            }

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

        public int buscarPalabrasClaves(string DigitarTexto, string palabras)
        {
            palabras = String.Format(palabras, "{0:D15}");

            return new Regex(@"(?:" + palabras + " )").Matches(DigitarTexto).Count;

        }

        public List<DomainValue> listadoValoresDominios()
        {
            var container = ODataContextWrapper.GetContainer();

            return container.DomainValues.Where(a => a.DomainId == 1009).ToList();
        }

        static string decodeBase64(string sInput)
        {
            String codedBody = Regex.Replace(sInput, "([-])", "+");
            codedBody = Regex.Replace(codedBody, "([_])", "+");
            byte[] data = Convert.FromBase64String(Regex.Replace(codedBody, "=", "/"));
            return Encoding.UTF8.GetString(data);
        }

        public void EvaluarPuntuacion(string asunto)
        {

            List<DomainValue> listado = listadoValoresDominios();


            string palabraClave = String.Empty;


            foreach (DomainValue b in listado)
            {
                palabraClave = b.Value.ToString();
                int contador = buscarPalabrasClaves(asunto, palabraClave);


                if (contador > 0)
                    puntos.Add(new Puntuacion() { contador = contador, palabra = palabraClave });
            }

            int[] valores = MaioresValores();

            //isto sera hecho depues de validar lo cuerpo tambien
            if (((valores[0] * 100) / (valores[0] + valores[1])) >= 70)
            {
                //Enviar a lo Crear Ticket Hijo con los parametros do workflow de la _palabra[0]; en ticketsvalues
            }
            else
            {

                // Enviar a la pantalla de classificaicon manual
            }
        }

        public int[] MaioresValores()
        {
            int[] valores = new int[2];

            foreach (Puntuacion p in puntos)
            {
                if (p.contador > valores[0])
                {
                    valores[1] = valores[0];
                    _palabras[1] = _palabras[0];
                    valores[0] = p.contador;
                    _palabras[0] = p.palabra;

                }
                else
                 if (p.contador > valores[1])
                {
                    valores[1] = p.contador;
                    _palabras[1] = p.palabra;
                }
            }

            return valores;

        }



    }
}