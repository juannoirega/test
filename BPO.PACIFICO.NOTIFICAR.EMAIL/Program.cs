using BPO.PACIFICO.NOTIFICAR.EMAIL;
using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services;
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

        static string[] _valores = new string[15];
        static string[] _valoresTicket = new string[6];
        string _contenido = String.Empty;
        string _correoRobot = "soportecorredor_des@pacifico.com.pe";
        string _rutaPlantilla = String.Empty;
        static string[] Scopes = { GmailService.Scope.GmailModify };
        static string ApplicationName = "Gmail API .NET Quickstart";
        UserCredential credential;
        static List<PalabrasClave> palabrasClaves = new List<PalabrasClave>();
        static JsonCorreo JsonCorreo = new JsonCorreo();
        Ticket ticket = new Ticket();
        List<TicketValue> ticketValue = new List<TicketValue>();
        List<TicketValue> ticketValue_Valores = new List<TicketValue>();
        DomainValue TipoProceso = null;

        static void Main(string[] args)
        {
            _robot = new BaseRobot<Program>(args);
            _robot.Start();
        }

        protected override void Start()
        {

            if (_robot.Tickets.Count < 1)
                return;

            LogStartStep(2);
            GetRobotParam();

            foreach (Ticket ticket in _robot.Tickets)
            {
                try
                {
                    ProcesarTicket(ticket);
                }
                catch (Exception ex)
                {
                    LogFailStep(30, ex);
                    cambiarEstado(TipoProceso.Value, ticket, "si", ex.Message);
                }

            }
        }

        public void ProcesarTicket(Ticket ticket)
        {
            ticket = _robot.Tickets.FirstOrDefault();



            _valoresTicket[0] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre_contratante).Value;
            _valoresTicket[1] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.numero_de_poliza).Value;
            //VALOR QUE DETERMINARA DEL DOMINIO FUNCIONAL Y QUE PLATILLA USAR PARA LA NOTIFICIACIÓN 
            _valoresTicket[2] = "1037";

            //Correos 
            _valoresTicket[3] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.email_solicitante).Value;
            //Correos Copias
            _valoresTicket[4] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.email_en_copia).Value;


            //Opteniendo el Nombre del Dominio Funcionanal para pasar al Siguiente Estado
            var container = ODataContextWrapper.GetContainer();
            TipoProceso = container.DomainValues.Where(c => c.Id == Convert.ToInt32(_valoresTicket[2])).FirstOrDefault();


            //Optener todos lo Ticket del Workflow "Adjuntar Documentos"
            List<Ticket> ticketDocumento = _robot.GetDataQueryTicket().Where(t => t.StateId == Convert.ToInt32(_valores[3])).ToList();

            //Almacenar los TicketValue de Tickets
            foreach (var item in ticketDocumento)
            {
                ticketValue.Add(_robot.GetDataQueryTicketValue().Where(t => t.TicketId == item.Id).OrderBy(t => t.Id).First());
                ticketValue.Add(_robot.GetDataQueryTicketValue().Where(t => t.TicketId == item.Id).OrderByDescending(t => t.Id).First());
            }

            //Optener el Ultimo Tickets Adjuntar Documentos
            var idticket = ticketValue.Where(t => t.Value == _valoresTicket[2]).OrderByDescending(t => t.Id).Select(t => new { t.TicketId }).FirstOrDefault();

            //Optener la Plantilla del ticketValue
            var valorTicket = ticketValue.Where(t => t.TicketId == Convert.ToInt32(idticket.TicketId) && t.FieldId == Convert.ToInt32(_valores[2])).OrderByDescending(t => t.Id).Select(t => new { t.Value }).FirstOrDefault();

            //Opteneniendo la Ruta de la Plantilla
            _rutaPlantilla = Convert.ToString(valorTicket.Value);

            LeerArchivo(_rutaPlantilla);

            //Metodo Email
            EnviarEmail();

           


            cambiarEstado(TipoProceso.Value, ticket,"no","");

        }

        public void GetRobotParam()
        {
            _valores[0] = _robot.GetValueParamRobot("EstadoSolicitudAceptadaA").ValueParam;
            _valores[1] = _robot.GetValueParamRobot("EstadoSolicitudRechazadaA").ValueParam;

            _valores[2] = _robot.GetValueParamRobot("FieldAdjuntarDocumentos").ValueParam;
            _valores[3] = _robot.GetValueParamRobot("EstadoAdjuntarDocumentos").ValueParam;
            _valores[4] = _robot.GetValueParamRobot("RutaArchivosPlantillas").ValueParam;
         

            _valores[6] = _robot.GetValueParamRobot("EstadoErrorMA").ValueParam;
            _valores[7] = _robot.GetValueParamRobot("EstadoErrorMR").ValueParam;

            _valores[8] = _robot.GetValueParamRobot("EstadoSolicitudAceptadaR").ValueParam;
            _valores[9] = _robot.GetValueParamRobot("EstadoSolicitudRechazadaR").ValueParam;

            _valores[10] = _robot.GetValueParamRobot("EstadoSolicitudAceptadaAC").ValueParam;
            _valores[11] = _robot.GetValueParamRobot("EstadoSolicitudRechazadaAC").ValueParam;

            _valores[12] = _robot.GetValueParamRobot("EstadoErrorMAC").ValueParam;

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
            //GetEncoding para ACeptar tilde y caracteres Especiales
            mail.BodyEncoding = Encoding.GetEncoding("iso-8859-1");
            mail.SubjectEncoding = Encoding.GetEncoding("iso-8859-1");
            mail.From = new MailAddress(_correoRobot);
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
            service.Users.Messages.Send(message, _correoRobot).Execute();
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

        public void cambiarEstado(String TipoProceso, Ticket ticket,String error,String mensaje)
        {
            switch (TipoProceso)
            {
                case "Plantilla Conforme Anulación Póliza":
                    if(error=="si")
                        _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, mensaje), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId ==  Convert.ToInt32(_valores[6])).Id);
                    else
                    _robot.SaveTicketNextState(ticket, Convert.ToInt32(_valores[0]));
                    break;
                case "Plantilla Rechazo Anulación Póliza":
                    if (error == "si")
                        _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, mensaje), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == Convert.ToInt32(_valores[6])).Id);
                    else
                        _robot.SaveTicketNextState(ticket, Convert.ToInt32(_valores[1]));
                    break;
                case "Plantilla Conforme Rehabilitación":
                    if (error == "si")
                        _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, mensaje), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == Convert.ToInt32(_valores[7])).Id);
                    else
                        _robot.SaveTicketNextState(ticket, Convert.ToInt32(_valores[8]));
                    break;
                case "Plantilla Rechazo Rehabilitación":
                    if (error == "si")
                        _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, mensaje), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == Convert.ToInt32(_valores[7])).Id);
                    else
                        _robot.SaveTicketNextState(ticket, Convert.ToInt32(_valores[9]));
                    break;
                case "Plantilla Conforme Actualización Cliente":
                    if (error == "si")
                        _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, mensaje), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == Convert.ToInt32(_valores[12])).Id);
                    else
                        _robot.SaveTicketNextState(ticket, Convert.ToInt32(_valores[10]));
                    break;
                case "Plantilla Rechazo Actualización CLiente":
                    if (error == "si")
                        _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, mensaje), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == Convert.ToInt32(_valores[12])).Id);
                    else
                        _robot.SaveTicketNextState(ticket, Convert.ToInt32(_valores[11]));
                    break;

            }

        }
    }
}
