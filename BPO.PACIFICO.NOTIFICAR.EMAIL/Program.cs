using BPO.PACIFICO.NOTIFICAR.EMAIL;
using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using Robot.Util.Nacar;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace BPO.Robot.Template.v3 //BPO.PACIFICO.NOTIFICAR.EMAIL
{
     class Program : IRobot
    {
        static BaseRobot<Program> _robot = null;

        private static Functions _Funciones;

        static string[] _valores = new string[10];
        static string[] _valoresTicket = new string[6];
        string _contenido = String.Empty;
        string _correoRobot = "bponaa@gmail.com";
        string _rutaPlantilla = String.Empty;
        static string[] Scopes = { GmailService.Scope.GmailModify };
        static string ApplicationName = "Gmail API .NET Quickstart";
        UserCredential credential;
        static List<PalabrasClave> palabrasClaves = new List<PalabrasClave>();
        static JsonCorreo JsonCorreo = new JsonCorreo();
        Ticket ticket = new Ticket();
        List<TicketValue> ticketValue = new List<TicketValue>();
        List<TicketValue> ticketValue_Valores = new List<TicketValue>();

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
            try
            {
                ticket = _robot.Tickets.FirstOrDefault();


                _valoresTicket[0] = "Luis Kevin Trujillo Hoyos";
                _valoresTicket[1] = "N° 12345678900";
                //VALOR QUE DETERMINARA DEL DOMINIO FUNCIONAL Y QUE PLANTILLA USARA PARA LA NOTIFICIACIÓN EMAIL
                _valoresTicket[2] = "8";
                //Correos 
                _valoresTicket[3] = "luistrujilloh@hotmail.com,ltrujill@everis.com";
                //Correos Copias
                _valoresTicket[4] = "bponaa@gmail.com";

                GetRobotParam();

                //Optener todos lo Ticket del Workflow "Adjuntar Documentos"
                List<Ticket> ticketDocumento = _robot.GetDataQueryTicket().Where(t => t.StateId == Convert.ToInt32(_valores[3])).ToList();

                //Almacenar los TicketValue de Tickets
                foreach (var item in ticketDocumento)
                {
                    ticketValue.Add(_robot.GetDataQueryTicketValue().Where(t => t.TicketId == item.Id).OrderBy(t => t.Id).First());
                    ticketValue.Add(_robot.GetDataQueryTicketValue().Where(t => t.TicketId == item.Id).OrderByDescending(t => t.Id).First());
                }

                //Opteneniendo la Ruta de la Plantilla
                _rutaPlantilla = ticketValue.OrderByDescending(t => t.Id).FirstOrDefault(o => o.TicketId == ticketValue.OrderByDescending(t => t.Id).FirstOrDefault(tv => tv.Value == _valoresTicket[2]).TicketId).Value;

                LeerArchivo(_rutaPlantilla);
                //Metodo Email
                EnviarEmail();

                if (_valoresTicket[2] == _valoresTicket[2])
                    _robot.SaveTicketNextState(ticket, Convert.ToInt32(_valores[0]));

                else
                    _robot.SaveTicketNextState(ticket, Convert.ToInt32(_valores[1]));
            }
            catch (Exception ex)
            {

                LogFailStep(30, ex);
                //_robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, ex.Message), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == _estadoError).Id);
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


        public void EnviarEmail()
        {

            AutenticacionEmail();
            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //Gmail API credentials
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

            //Create Message
            MailMessage mail = new MailMessage();
            mail.Subject = JsonCorreo.Subject;
            mail.Body = JsonCorreo.Body;
            mail.From = new MailAddress(_correoRobot);
            mail.IsBodyHtml = true;
            //Adjuntar Documentos
            //string docume = @"C:\Users\ltrujill\Documents\JSON.txt";
            //mail.Attachments.Add(new Attachment(docume));
            
            ////Un Correo
            //mail.To.Add(new MailAddress(_correoRobot));
     
            mail.To.Add(FormatMultipleEmailAddresses(_valoresTicket[3]));
            mail.CC.Add(FormatMultipleEmailAddresses(_valoresTicket[4]));
            MimeKit.MimeMessage mimeMessage = MimeKit.MimeMessage.CreateFromMailMessage(mail);

            Message message = new Message();
            message.Raw = Base64UrlEncode(mimeMessage.ToString());

            //Send Email
            var result = service.Users.Messages.Send(message, _correoRobot).Execute();
        }

        private string FormatMultipleEmailAddresses(string emailAddresses)
        {
            return string.Join(",", emailAddresses.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries));
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
        
        public void LeerArchivo(string archivo)
        {

            try
            {
                //Lleyendo el JSON
                using (StreamReader lector = new StreamReader(String.Concat(@"",_valores[4],archivo,"")))
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

                //Deserializando JSON en una CLase
                JsonCorreo = JsonConvert.DeserializeObject<JsonCorreo>(_contenido);

                //Almacenando Palabras que esten entre {XXX} y asignando la palabras para reemplazar
                foreach (Match match in Regex.Matches(JsonCorreo.Body, @"\{([^{}\]]*)\}"))
                    if (match.Value.Length >= 11)
                        palabrasClaves.Add(new PalabrasClave() { clave = match.Value, palabra = _valoresTicket[1] });
                    else
                        palabrasClaves.Add(new PalabrasClave() { clave = match.Value, palabra = _valoresTicket[0] });
                
                
                //Reemplazar las palabras Claves para enviar Correos
                foreach (PalabrasClave p in palabrasClaves)
                    JsonCorreo.Body = ReemplazarPalabras(JsonCorreo.Body, p.clave, p.palabra);

            }

            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

        }

        private string Base64UrlEncode(string input)
        {
            var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(inputBytes)
              .Replace('+', '-')
              .Replace('/', '_')
              .Replace("=", "");
        }

        public String ReemplazarPalabras(String texto, String palabra, String reemplazar)
        {
            return Regex.Replace(texto, @"(?:" + palabra + ")", "" + reemplazar + "");
        }
    }
}
