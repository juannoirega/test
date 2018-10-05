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
        static string[] _valores = new string[12];
        static string _userId = "soportecorredor_des@pacifico.com.pe";
        static List<string> _adjuntos = null;
        static int[] _fields = { eesFields.Default.email_cuerpo, eesFields.Default.email_asunto, eesFields.Default.error_des, eesFields.Default.id_est_hijo, eesFields.Default.id_est_padre, eesFields.Default.fields, eesFields.Default.email_fecha_hora, eesFields.Default.email_de, eesFields.Default.email_cc, eesFields.Default.id_est_mesa_control, eesFields.Default.fec_solicitud, eesFields.Default.idproceso };
        static string ApplicationName = "Gmail API .NET Quickstart";
        static List<DomainValue> _listado = null;
        static List<Puntuacion> puntos = new List<Puntuacion>();
        static string _diretorio = String.Empty;
        static string _domainpalabrasclaves = String.Empty;
        #endregion
        static void Main(string[] args)
        {
            try
            {
                _robot = new BaseRobot<Program>(args);
                _robot.Start();
            }
            catch (Exception Ex) { _robot.Start();  Console.WriteLine(Ex.Message); }
        }

        protected override void Start()
        {
            try
            {
                LogStartStep(2);
                GetRobotParam();
                Inicio();
                Email();
            }
            catch (Exception ex)
            {
                LogFailStep(30, ex);
            }
            finally
            {
                Environment.Exit(0);
            }
        }

        private void Inicio()
        {
            Console.WriteLine("♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦ ROBOT ♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦");
            Console.WriteLine("                Robot Captura Email                    ");
            Console.WriteLine("♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦");
        }

        public void GetRobotParam()
        {
            _valores[9] = _robot.GetValueParamRobot("EstadoError").ValueParam;
            _valores[3] = _robot.GetValueParamRobot("EstadoHijo").ValueParam;
            _valores[4] = _robot.GetValueParamRobot("EstadoPadre").ValueParam;
            _domainpalabrasclaves = _robot.GetValueParamRobot("DominioPalabrasClaves").ValueParam;
            _diretorio = _robot.GetValueParamRobot("Diretorio").ValueParam;
            LogEndStep(4);
        }

        public void Email()
        {

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.ReadWrite))
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
                _listado = ListadoValoresDominios();
            }
            catch (Exception ex) { throw new Exception("No se pudo buscar los dominos funcionales" + ex.Message); }
            
            // List Messages.
            IList<Message> messages = request.Execute().Messages;
            IList<String> labelsRemove = new String[] { "UNREAD" };

            if (messages != null && messages.Count > 0)
            {
                foreach (Message message in messages)
                {
                    Message infoResponse = service.Users.Messages.Get(_userId, message.Id).Execute();
                    AcionRequest(infoResponse, service);

                    //Marcar correo como leído:
                   // ModifyThread(service, message.ThreadId, message.LabelIds,labelsRemove);

                    _listado.Clear();
                }
            }
            else
            {
                LogInfoStep(44);
            }
        }

        //Modifica la etiqueta de correo como LEÍDO:
        public static Google.Apis.Gmail.v1.Data.Thread ModifyThread(GmailService oService, string threadId, IList<String> labelsToAdd, IList<String> labelsToRemove)
        {
            ModifyThreadRequest modify = new ModifyThreadRequest();
            modify.AddLabelIds = labelsToAdd;
            modify.RemoveLabelIds = labelsToRemove;

            try
            {
                return oService.Users.Threads.Modify(modify, _userId, threadId).Execute();
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error: " + Ex.Message, Ex); }
        }

        //cada message tiene una acion de requerimento
        public void AcionRequest(Message infoResponse, GmailService service)
        {
            LogStartStep(50);

            if (infoResponse != null)
            {
                String body = String.Empty;
                try
                {
                    _valores[6] = infoResponse.Payload.Headers.FirstOrDefault(o => o.Name == "Date").Value;
                    _valores[10] = Convert.ToDateTime(infoResponse.Payload.Headers.FirstOrDefault(o => o.Name == "Date").Value).ToShortDateString();
                    _valores[7] = infoResponse.Payload.Headers.FirstOrDefault(o => o.Name == "From").Value;
                    _valores[1] = infoResponse.Payload.Headers.FirstOrDefault(o => o.Name == "Subject").Value;

                    if (infoResponse.Payload.Headers.FirstOrDefault(o => o.Name == "Cc") != null)
                        _valores[8] = infoResponse.Payload.Headers.FirstOrDefault(o => o.Name == "Cc").Value;

                }
                catch (Exception ex) { throw new Exception("No se pudo guardar el encabezado de correo electrónico" + ex.Message); }

                if (_valores[6] != "" && _valores[7] != "")
                {
                    if (infoResponse.Payload.Parts == null && infoResponse.Payload.Body != null)
                    {
                        body = infoResponse.Payload.Body.Data;
                    }
                    else
                    {
                        LogStartStep(52);
                        try
                        {
                            body = GetNestedParts(infoResponse.Payload.Parts, "");
                        }
                        catch (Exception ex) { throw new Exception("No se pudo guardar el cuerpo del email" + ex.Message); }
                    }

                    if (infoResponse.Payload.Parts != null)
                        if (infoResponse.Payload.Parts.FirstOrDefault(o => o.Filename != "") != null)
                        {
                            _adjuntos = GetFiles(service, infoResponse.Payload.Parts, infoResponse.Id);
                        }

                    _valores[0] = DecodeBase64(body);
                    EvaluarPuntuacion(String.Concat(DecodeBase64(body), " ", _valores[1]), new Ticket { Priority = PriorityType.Media, RobotVirtualMachineId = null, StateId = null });
                    Array.Clear(_valores, 0, _valores.Length);
                }
            }
        }

        // cada acion de requerimento pueede tener o no Adjuntos e hacer lo getfiles
        private List<string> GetFiles(GmailService service, IList<MessagePart> parts, string messageId)
        {
            LogStartStep(51);
            try
            {
                List<string> files = new List<string>();
                foreach (MessagePart part in parts)
                {

                    if (!String.IsNullOrEmpty(part.Filename) && part.Filename.Substring(part.Filename.Length - 3, 3) != "gif")
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
            }
            catch (Exception ex) { throw new Exception("No fue posible guardar los adjuntos" + ex.Message); }
        }

        static String GetNestedParts(IList<MessagePart> part, string curr)
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
                        return GetNestedParts(parts.Parts, str);
                    }
                }

                return str;
            }

        }

        public int BuscarPalabrasClaves(string DigitarTexto, string palabras)
        {
            LogStartStep(54);
            try
            {
                palabras = String.Format(palabras, "{0:D15}");

                return new Regex(String.Concat(@"(?:", palabras, " )")).Matches(DigitarTexto.ToUpperInvariant()).Count;
            }
            catch (Exception ex) { throw new Exception("No se pudo encontrar las palabras clave" + ex.Message); }

        }

        public List<DomainValue> ListadoValoresDominios()
        {
            var container = ODataContextWrapper.GetContainer();

            return container.DomainValues.Where(a => a.DomainId ==Convert.ToInt32(_domainpalabrasclaves)).ToList();
        }

        static string DecodeBase64(string sInput)
        {
            try
            {
                String codedBody = Regex.Replace(sInput, "([-])", "+");
                codedBody = Regex.Replace(codedBody, "([_])", "+");
                byte[] data = Convert.FromBase64String(Regex.Replace(codedBody, "=", "/"));
                string texto = Encoding.UTF8.GetString(data);
                var match = Regex.Match(texto, "(<div\\s)");
                if (match.Success)
                    return texto.Substring(match.Index, texto.Length - match.Index);
                else
                    return texto;
            }
            catch (Exception ex) { throw new Exception("No se pudo convertir la Base64" + ex.Message); }
        }

        // cada acion de erquerimento tiene una evaluacion de puentos 
        public void EvaluarPuntuacion(string texto, Ticket ticket)
        {
            LogStartStep(53);

            try
            {
                string palabraClave = String.Empty;

                foreach (DomainValue b in _listado)
                {
                    palabraClave = b.Value.ToString().ToUpperInvariant();
                    int contador = BuscarPalabrasClaves(texto, palabraClave);

                    if (contador > 0)
                        puntos.Add(new Puntuacion() { contador = contador, palabra = palabraClave });
                }

                int[] valores = MaioresValores();

                if (AdicionarNumeroPolizaoDnioRuc(ticket, texto) && valores[0] > 0 && ((valores[0] * 100) / (valores[0] + valores[1])) >= 70)
                {

                    CreacionTicket(texto, ticket, true);
                }
                else
                {
                    CreacionTicket(texto, ticket, false);
                }
            }
            catch (Exception ex) { throw new Exception("No fue posible analizar la puntuación" + ex.Message); }
        }

        public void CreacionTicket(string texto, Ticket ticketPadre, bool flag)
        {
            LogStartStep(55);
            try
            {
                DatosFields();
                AdicionarValues(ticketPadre);

                if (_adjuntos != null)
                    AdicionarAdjuntos(ticketPadre);

                if (flag)
                    ticketPadre.StateId = 2;
                else
                    ticketPadre.StateId = 2;

                _robot.SaveNewTicket(ticketPadre);
            }
            catch (Exception ex) { throw new Exception("No se pudo crear el Ticket" + ex.Message); }
        }

        public void DatosFields()
        {
            foreach (int field in _fields)
                _valores[5] = String.Concat(field.ToString(), ",", _valores[5]);
        }

        public void AdicionarAdjuntos(Ticket ticket)
        {
            for (int num = 0; _adjuntos.Count > num; num++)
                ticket.TicketValues.Add(new TicketValue { Value = _adjuntos[num], ClonedValueOrder = num, TicketId = ticket.Id, FieldId = eesFields.Default.endoso_adj });
        }

        public void AdicionarValues(Ticket ticket)
        {
            for (int cont = 0; _fields.Length > cont; cont++)
                ticket.TicketValues.Add(new TicketValue { Value = _valores[cont], ClonedValueOrder = null, TicketId = ticket.Id, FieldId = _fields[cont] });

            ticket.TicketValues.Add(new TicketValue { Value = _userId, ClonedValueOrder = null, TicketId = ticket.Id, FieldId = eesFields.Default.email_para });
        }

        public bool AdicionarNumeroPolizaoDnioRuc(Ticket ticket, string texto)
        {
            string police = Regex.Match(texto, @"[^1-9]((2[1-9])[0-9]{4}[0-9]{4})([^1-9]|$)").Value;
            string dni = Regex.Match(texto, @"[^1-9](([1-9])[0-9]{4}[0-9]{4})([^1-9]|$)").Value;
            string ruc = Regex.Match(texto, @"[^1-9](([1-9])[0-9]{5}[0-9]{5})([^1-9]|$)").Value;
            ruc = ValidarRuc(ruc);
            if (!String.IsNullOrWhiteSpace(police))
            {
                ticket.TicketValues.Add(new TicketValue { Value = police.Substring(1, 10), ClonedValueOrder = null, TicketId = ticket.Id, FieldId = eesFields.Default.poliza_nro });
                _valores[5] = _valores[5] + eesFields.Default.poliza_nro + ",";
                return true;
            }
            if (!String.IsNullOrWhiteSpace(dni))
            {
                ticket.TicketValues.Add(new TicketValue { Value = dni.Substring(1, 9), ClonedValueOrder = null, TicketId = ticket.Id, FieldId = eesFields.Default.nro_dni  });
                _valores[5] = _valores[5] + eesFields.Default.nro_dni + ",";
                return true;
            }
            else if (!String.IsNullOrWhiteSpace(ruc))
            {
                ticket.TicketValues.Add(new TicketValue { Value = ruc.Substring(1, 11), ClonedValueOrder = null, TicketId = ticket.Id, FieldId = eesFields.Default.nro_ruc  });
                _valores[5] = _valores[5] + eesFields.Default.nro_ruc + ",";
                return true;

            }

            return false;
        }

        public string ValidarRuc(string ruc)
        {
            if (String.IsNullOrWhiteSpace(ruc))
                return null;

            char[] numeros = ruc.Substring(1, 11).ToCharArray();

            int total = 0;
            int[] operador = new int[] { 5, 4, 3, 2, 7, 6, 5, 4, 3, 2 };

            for (int i = 0; i < 10; i++)
                total += Convert.ToInt32(numeros[i]) * operador[i];

            if ((total % 11) == Convert.ToInt32(numeros[10]))
                return ruc;
            else
                return null;

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