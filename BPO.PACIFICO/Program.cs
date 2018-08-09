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
using Everis.Ees.Entities.Enums;
using Robot.Util.Nacar;

namespace GmailQuickstart
{
    class Program : IRobot
    {
        #region Paremetros
        static BaseRobot<Program> _robot = null;
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/gmail-dotnet-quickstart.json
        static string[] Scopes = { GmailService.Scope.GmailReadonly };
        static string[] _palabras = new string[2];
        static string[] _valores = new string[10];
        static string _userId = "bponaa@gmail.com";
        static List<string> _adjuntos = null;
        static int[] _fields = { eesFields.Default.cuerpo_de_email, eesFields.Default.asunto_de_email, eesFields.Default.estado_error, eesFields.Default.estado_hijo, eesFields.Default.estado_padre, eesFields.Default.fields, eesFields.Default.fecha_hora_de_email };
        static string ApplicationName = "Gmail API .NET Quickstart";
        static List<DomainValue> _listado = null;
        static List<Puntuacion> puntos = new List<Puntuacion>();
        static string _diretorio = String.Empty;
        #endregion
        static void Main(string[] args)
        {
            _robot = new BaseRobot<Program>(args);
            _robot.Start();

        }

        protected override void Start()
        {
            //    GetRobotParam();
            Email();
        }

        public void GetRobotParam()
        {
            _valores[2] = _robot.GetValueParamRobot("EstadoError").ValueParam;
            _valores[3] = _robot.GetValueParamRobot("EstadoHijo").ValueParam;
            _valores[4] = _robot.GetValueParamRobot("EstadoPadre").ValueParam;

            _diretorio = _robot.GetValueParamRobot("Diretorio").ValueParam;
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
                }

                // Create Gmail API service.
                GmailService service = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                // Define parameters of request.
                UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(_userId);
                request.MaxResults = 5;
                request.LabelIds = "INBOX";
                request.IncludeSpamTrash = false;
                request.Q = "is:unread";
                try
                {
                    _listado = listadoValoresDominios();
                }
                catch(Exception ex){ throw new Exception(); }
                // List Messages.
                IList<Message> messages = request.Execute().Messages;

                if (messages != null && messages.Count > 0)
                {
                    foreach (Message message in messages)
                    {
                        Message infoResponse = service.Users.Messages.Get(_userId, message.Id).Execute();
                        AcionRequest(infoResponse, service);
                        _listado.Clear();
                    }
                }
                else
                {
                    Console.WriteLine("No messages found.");
                }
            


        }
        //cada message tiene una acion de requerimento
        public void AcionRequest(Message infoResponse, GmailService service)
        {
            

            if (infoResponse != null)
            {
                String body = String.Empty;
                try
                {
                    _valores[6] = infoResponse.Payload.Headers.FirstOrDefault(o => o.Name == "Date").Value;

                    _valores[7] = infoResponse.Payload.Headers.FirstOrDefault(o => o.Name == "From").Value;

                    _valores[1] = infoResponse.Payload.Headers.FirstOrDefault(o => o.Name == "Subject").Value;

                    _valores[8] = infoResponse.Payload.Headers.FirstOrDefault(o => o.Name == "Cc").Value;

                }catch(Exception ex) { }

                    if (_valores[6] != "" && _valores[7] != "")
                    {
                        if (infoResponse.Payload.Parts == null && infoResponse.Payload.Body != null)
                        {
                            body = infoResponse.Payload.Body.Data;
                        }
                        else
                        {
                            body = getNestedParts(infoResponse.Payload.Parts, "");
                        }
                  
                        if (infoResponse.Payload.Parts.FirstOrDefault(o => o.Filename != "")!=null)
                        {
                            _adjuntos = GetFiles(service, infoResponse.Payload.Parts, infoResponse.Id);
                        }
                   
                 

                        _valores[0] = decodeBase64(body);


                        EvaluarPuntuacion(String.Concat(decodeBase64(body), " ", _valores[1]), new Ticket { Priority = PriorityType.Media, RobotVirtualMachineId = null, StateId = null });

                        Array.Clear(_valores, 0, _valores.Length);

                    }

                
            }

        }
        // cada acion de requerimento pueede tener o no Adjuntos e hacer lo getfiles
        private List<string> GetFiles(GmailService service, IList<MessagePart> parts, string messageId)
        {
            try
            {
                List<string> files = new List<string>();
                foreach (MessagePart part in parts)
                {
                    if (!String.IsNullOrEmpty(part.Filename))
                    {
                        String attId = part.Body.AttachmentId;

                        MessagePartBody attachPart = service.Users.Messages.Attachments.Get(_userId, messageId, attId).Execute();
                        String attachData = Regex.Replace(attachPart.Data, "-", "+");
                        attachData = Regex.Replace(attachData, "_", "/");
                        attachData = Regex.Replace(attachData, "=", "/");
                        byte[] data = Convert.FromBase64String(attachData);
                        File.WriteAllBytes(Path.Combine(_diretorio, part.Filename), data);
                        files.Add(Path.Combine(_diretorio, part.Filename));
                    }
                }
                return files;
            }catch(Exception ex) { throw new Exception(); }


        }

        static String getNestedParts(IList<MessagePart> part, string curr)
        {
            try
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
            }catch(Exception ex) { throw new Exception(); }

        }

        public int buscarPalabrasClaves(string DigitarTexto, string palabras)
        {
            try
            {
                palabras = String.Format(palabras, "{0:D15}");

                return new Regex(String.Concat(@"(?:", palabras, " )")).Matches(DigitarTexto).Count;
            }catch(Exception ex) { throw new Exception(); }

        }

        public List<DomainValue> listadoValoresDominios()
        {
            var container = ODataContextWrapper.GetContainer();

            return container.DomainValues.Where(a => a.DomainId == 1009).ToList();

        }


        static string decodeBase64(string sInput)
        {
            try
            {
                String codedBody = Regex.Replace(sInput, "([-])", "+");
                codedBody = Regex.Replace(codedBody, "([_])", "+");
                byte[] data = Convert.FromBase64String(Regex.Replace(codedBody, "=", "/"));
                return Encoding.UTF8.GetString(data);
            }catch(Exception ex) { throw new Exception(); }
        }

        // cada acion de erquerimento tiene una evaluacion de puentos 
        public void EvaluarPuntuacion(string texto, Ticket ticket)
        {
            try
            {


                string palabraClave = String.Empty;

                foreach (DomainValue b in _listado)
                {
                    palabraClave = b.Value.ToString();
                    int contador = buscarPalabrasClaves(texto, palabraClave);


                    if (contador > 0)
                        puntos.Add(new Puntuacion() { contador = contador, palabra = palabraClave });
                }

                int[] valores = MaioresValores();


                if (((valores[0] * 100) / (valores[0] + valores[1])) >= 70)
                {

                    CreacionTicket(texto, ticket, true);
                }
                else
                {

                    CreacionTicket(texto, ticket, false);
                }
            }catch(Exception ex) { }
        }


        public void CreacionTicket(string texto, Ticket ticketPadre, bool flag)
        {
            try
            {
                DatosFields();
                AdicionarNumeroPoliza(ticketPadre, texto);
                AdicionarValues(ticketPadre);

                if (_adjuntos != null)
                    AdicionarAdjuntos(ticketPadre);

                if (flag)
                    ticketPadre.StateId = 4;
                else
                    ticketPadre.StateId = 3;

                _robot.SaveNewTicket(ticketPadre);
            }
            catch(Exception ex) { throw new Exception(); }

        }
        public void DatosFields()
        {

            foreach (int field in _fields)
                _valores[5] = String.Concat(field.ToString(), ",", _valores[5]);
        }
        public void AdicionarAdjuntos(Ticket ticket)
        {
            string value = String.Empty;

            foreach (string doc in _adjuntos)
                value = String.Concat(doc, ",");

            ticket.TicketValues.Add(new TicketValue { Value = value, ClonedValueOrder = null, TicketId = ticket.Id, FieldId = eesFields.Default.documentos });
        }

        public void AdicionarValues(Ticket ticket)
        {
            for (int cont = 0; 5 >= cont; cont++)
                ticket.TicketValues.Add(new TicketValue { Value = _valores[cont], ClonedValueOrder = null, TicketId = ticket.Id, FieldId = _fields[cont] });

        }

        public void AdicionarNumeroPoliza(Ticket ticket, string texto)
        {
            ticket.TicketValues.Add(new TicketValue { Value = Regex.Match(texto, "(2[1-9])[0-9]{4}[0-9]{4}").Value, ClonedValueOrder = null, TicketId = ticket.Id, FieldId = eesFields.Default.numero_de_poliza });

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