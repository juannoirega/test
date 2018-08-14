using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace BPO.PACIFICO.NOTIFICACION.EMAIL
{
    class Program : IRobot
    {
        static BaseRobot<Program> _robot = null;

        static string[] _valores = new string[10];
        static string[] _valoresTicket = new string[4];
        string _contenido = String.Empty;
        string _correoRobot = "bponaa@gmail.com";
        string _asunto = String.Empty;
        string _cuerpo = String.Empty;
        Int32 idAccion = 0;
        static string[] Scopes = { GmailService.Scope.GmailSend };
        static string ApplicationName = "Gmail API .NET Quickstart";
        UserCredential credential;
        static List<PalabrasClave> palabrasClaves = new List<PalabrasClave>();
        static JsonCorreo JsonCorreo = new JsonCorreo();
        Ticket ticket = new Ticket();
        TicketValue ticketValue = new TicketValue();

        static void Main(string[] args)
        {
            _robot = new BaseRobot<Program>(args);
            _robot.Start();
        }

        protected override void Start()
        {
            ProcesarTicket();
        }

        public void ProcesarTicket()
        {
            ticket = _robot.Tickets.FirstOrDefault();

            //Datos del Ticket
            //_valoresTicket[0] = ticket.TicketValues[0].Value;
            //_valoresTicket[1] = ticket.TicketValues[1].Value;
            //_valoresTicket[1] = ticket.TicketValues[1].Value;

            _valoresTicket[0] = "Luis Kevin Trujillo Hoyos";
            _valoresTicket[1] = "N° 12345678900";
            _valoresTicket[2] = "bponaa@gmail.com";

            //VALOR QUE DETERMINARA A QUE ESTADO DIRIJIRA EL ROBOT NOTIFICIÓN EMAIL
            _valoresTicket[3] = "Solicitud Aceptada";


            GetRobotParam();

            //Optener el Ultimo ticket Workflow "Subir Documentos"   
            Ticket ticketDocumento = _robot.GetDataQueryTicket().Where(t => t.StateId == Convert.ToInt32(_valores[3])).OrderByDescending(t => t.Id).FirstOrDefault();
            //Optener el Ultimo TicketValue Workflow "Subir Documentos" 
            ticketValue = _robot.GetDataQueryTicketValue().Where(t => t.TicketId == ticketDocumento.Id && t.FieldId == Convert.ToInt32(_valores[2])).OrderByDescending(t => t.Id).FirstOrDefault();

            EnviarEmail();

            if (_valoresTicket[3] == "Solicitud Aceptada")
                _robot.SaveTicketNextState(ticket, Convert.ToInt32(_valores[0]));

            else
                _robot.SaveTicketNextState(ticket, Convert.ToInt32(_valores[1]));

        }



        public void EnviarEmail()
        {
            AutenticacionEmail();

            //// Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


            LeerArchivo(ticketValue.Value);


            string plainText = "To: " + _valoresTicket[2] + "\r\n" +
                                 "Subject: " + _asunto + "\r\n" +
                                 "Content-Type: text/html; charset=us-ascii\r\n\r\n" +
                                 JsonCorreo.Body;

            var newMsg = new Google.Apis.Gmail.v1.Data.Message();
            newMsg.Raw = Program.Base64UrlEncode(plainText.ToString());
            service.Users.Messages.Send(newMsg, _correoRobot).Execute();
        }

        public void AutenticacionEmail()
        {
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


        }

        public void GetRobotParam()
        {
            _valores[0] = _robot.GetValueParamRobot("EstadoSolicitudAceptada").ValueParam;
            _valores[1] = _robot.GetValueParamRobot("EstadoSolicitudRechazada").ValueParam;
            _valores[2] = _robot.GetValueParamRobot("FildAdjuntarDocuementos").ValueParam;
            _valores[3] = _robot.GetValueParamRobot("EstadoAdjuntarDocumentos").ValueParam;
            _valores[4] = _robot.GetValueParamRobot("RutaArchivosPlantillas").ValueParam;

        }

        public void LeerArchivo(string archivo)
        {

            string fichero = @"" + _valores[4] + archivo + "";

            try
            {
                using (StreamReader lector = new StreamReader(fichero))
                {
                    while (lector.Peek() > -1)
                    {
                        string linea = lector.ReadLine();
                        if (!String.IsNullOrEmpty(linea))
                        {
                            _contenido += linea;
                        }
                    }
                }

                JsonCorreo = JsonConvert.DeserializeObject<JsonCorreo>(_contenido);
                _asunto = JsonCorreo.Subject;
                //_cuerpo = JsonCorreo.Body;


                foreach (Match match in Regex.Matches(JsonCorreo.Body, @"\{([^{}\]]*)\}"))
                {
                    if (match.Value.Length >= 11)
                        palabrasClaves.Add(new PalabrasClave() { clave = match.Value, palabra = _valoresTicket[1] });
                    else
                        palabrasClaves.Add(new PalabrasClave() { clave = match.Value, palabra = _valoresTicket[0] });
                }



                foreach (PalabrasClave p in palabrasClaves)
                {

                    JsonCorreo.Body = ReemplazarPalabras(JsonCorreo.Body, p.clave, p.palabra);
                }

            }

            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

        }



        public static string Base64UrlEncode(string input)
        {
            var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(inputBytes).Replace("+", "-").Replace("/", "_").Replace("=", "");
        }
        public String ReemplazarPalabras(String texto, String palabra, String reemplazar)
        {
            return Regex.Replace(texto, @"(?:" + palabra + ")", "" + reemplazar + "");
        }

    }
}
