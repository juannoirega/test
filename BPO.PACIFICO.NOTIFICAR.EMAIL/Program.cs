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

        static string[] _valores = new string[20];
        static string[] _valoresTicket = new string[8];
        string _contenido = String.Empty;
        string _correoRobot = "soportecorredor_des@pacifico.com.pe";
        string _rutaPlantilla = String.Empty;
        string NombrePlantilla = String.Empty;
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
            _Funciones = new Functions();
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
                    EdentificarProceso(ticket, ex.Message);
                }

            }
        }

        public void ProcesarTicket(Ticket ticket)
        {
            ticket = _robot.Tickets.FirstOrDefault();


            //CAMPO QUE VIENE DE LOS ROBOT QUE INDICA QUE TIPO DE PLANTILLA USARA 
            _valoresTicket[2] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.id_archivo_tipo_adj).Value;
            //idProceso 
            _valoresTicket[7] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.idproceso).Value;
            //Asegurado Nombre 
            _valoresTicket[0] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.asegurado_nombre).Value;
            //Poliza
            _valoresTicket[1] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_nro).Value;
            //Correos 
            _valoresTicket[3] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.email_para).Value;
            //Correos Copias
            _valoresTicket[4] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.email_cc).Value;
            //Dni
            _valoresTicket[6] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nro_dni).Value;

            //Optenr el NOmbre de la Plantilla
            var dominio = ODataContextWrapper.GetContainer();
            NombrePlantilla = dominio.DomainValues.Where(c => c.Id == Convert.ToInt32(_valoresTicket[2])).FirstOrDefault().Value;

            //Optener todos lo Ticket del Workflow "Adjuntar Documentos"
            List<Ticket> ticketDocumento = _robot.GetDataQueryTicket().Where(t => t.StateId == Convert.ToInt32(_valores[3])).ToList();

            //Almacenar los TicketValue de Tickets
            foreach (var item in ticketDocumento)
            {
                ticketValue.Add(_robot.GetDataQueryTicketValue().Where(t => t.TicketId == item.Id).OrderBy(t => t.Id).First());
                ticketValue.Add(_robot.GetDataQueryTicketValue().Where(t => t.TicketId == item.Id).OrderByDescending(t => t.Id).First());
            }

            //Optener el Ultimo Tickets Adjuntar Documentos dependiendo el Tipo de Plantilla
            var idticket = ticketValue.Where(t => t.Value == _valoresTicket[2]).OrderByDescending(t => t.Id).Select(t => new { t.TicketId }).FirstOrDefault();

            //Optener la Plantilla del ticketValue
            var valorTicket = ticketValue.Where(t => t.TicketId == Convert.ToInt32(idticket.TicketId) && t.FieldId == Convert.ToInt32(_valores[2])).OrderByDescending(t => t.Id).Select(t => new { t.Value }).FirstOrDefault();

            //Opteneniendo la Ruta de la Plantilla
            _rutaPlantilla = Convert.ToString(valorTicket.Value);

            LeerArchivo(_rutaPlantilla,ticket);

            //Metodo Email
            EnviarEmail();

            EdentificarProceso(ticket, "");
           

        }

        public void GetRobotParam()
        {
            _valores[0] = _robot.GetValueParamRobot("EstadoSolicitudAceptadaAP").ValueParam;
            _valores[1] = _robot.GetValueParamRobot("EstadoSolicitudRechazadaAP").ValueParam;
            _valores[13] = _robot.GetValueParamRobot("MesaControlAP").ValueParam;

            _valores[2] = _robot.GetValueParamRobot("FieldAdjuntarDocumentos").ValueParam;
            _valores[3] = _robot.GetValueParamRobot("EstadoAdjuntarDocumentos").ValueParam;
            _valores[4] = _robot.GetValueParamRobot("RutaArchivosPlantillas").ValueParam;


            _valores[6] = _robot.GetValueParamRobot("EstadoErrorMA").ValueParam;
            _valores[7] = _robot.GetValueParamRobot("EstadoErrorMR").ValueParam;

            _valores[8] = _robot.GetValueParamRobot("EstadoSolicitudAceptadaRE").ValueParam;
            _valores[9] = _robot.GetValueParamRobot("EstadoSolicitudRechazadaRE").ValueParam;
            _valores[14] = _robot.GetValueParamRobot("MesaControlRE").ValueParam;


            _valores[10] = _robot.GetValueParamRobot("EstadoSolicitudAceptadaAC").ValueParam;
            _valores[11] = _robot.GetValueParamRobot("EstadoSolicitudRechazadaAC").ValueParam;
            _valores[15] = _robot.GetValueParamRobot("MesaControlAC").ValueParam;

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

            //GetEncoding para Aceptar tilde y caracteres Especiales
            mail.BodyEncoding = Encoding.GetEncoding("iso-8859-1");
            mail.SubjectEncoding = Encoding.GetEncoding("iso-8859-1");
            mail.From = new MailAddress(_correoRobot);

            //Adjuntar Documentos
            //List<string> ListaDocuemntos = new List<string>();
            //ListaDocuemntos.Add(@"\\PCLCEVE0Q6K\Content/Upload_Files/6/2018-09-06/3e96f340-289a-4f4c-bebc-f870894b8e3a_json_Anulacion_Conforme.TXT");
            //ListaDocuemntos.Add(@"\\PCLCEVE0Q6K\Content/Upload_Files/6/2018-09-06/3e96f340-289a-4f4c-bebc-f870894b8e3a_json_Anulacion_Conforme.TXT");

            //if (ListaDocuemntos != null)
            //{
            //    foreach (string archivos in ListaDocuemntos)
            //    {
            //        if (System.IO.File.Exists(archivos))
            //            mail.Attachments.Add(new Attachment(archivos));
            //    }
            //}



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

        public void LeerArchivo(string archivo,Ticket ticket)
        {
            String rutax = _valores[4] + archivo;
            rutax = rutax.Replace("/", "\\");

            try
            {
                //Lleyendo el JSON
                using (StreamReader lector = new StreamReader(@rutax))
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

                string valor = "";
                var containeValores = ODataContextWrapper.GetContainer();
                //Almacenando Palabras que esten entre {XXXXXX} y asignando la palabras para reemplazar
                foreach (Match match in Regex.Matches(JsonCorreo.Body, @"\{([^{}\]]*)\}"))
                {
                    valor = match.Value.Replace("{", "");
                    valor = valor.Replace("}", "");

                    //Opteniendo el Id del Field buscando por name
                    var fieldId = containeValores.Fields.Where(f => f.Name == valor).FirstOrDefault().Id;
                    //Opteniendo el Valor del ticker filtrandolo 
                    var name = ticket.TicketValues.Where(t => t.FieldId == fieldId).FirstOrDefault().Value;
                    valor = "{" + valor + "}";
                    palabrasClaves.Add(new PalabrasClave() { clave = valor, palabra = name });
                }



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

        public String buscarPalabras(String texto)
        {
            String palabras = String.Empty;

            String[] palabrasClavesPlantillas = ("CONFORME,RECHAZO").Split(',');
            int conta = 0;

            foreach (String item in palabrasClavesPlantillas)
            {
                palabras = item;
                palabras = String.Format(palabras, "{0:D15}");

                conta = new Regex(String.Concat(@"(?:"+palabras+")")).Matches(texto.ToUpperInvariant()).Count;
                if (conta == 1) palabras = item;
            }

            return palabras;
        }

        public void EdentificarProceso(Ticket ticket,string mensaje)
        {

            switch (Convert.ToInt32(_valoresTicket[7]))
            {
               case 1:
                    cambiarEstado(eesDomains.Default.AnulacionPoliza,buscarPalabras(NombrePlantilla),ticket,Convert.ToInt32(_valores[0]),Convert.ToInt32(_valores[1]),Convert.ToInt32(_valores[13]),mensaje);
                break;
                case  2:
                     cambiarEstado(eesDomains.Default.Rehabilitacion, buscarPalabras(NombrePlantilla), ticket, Convert.ToInt32(_valores[8]), Convert.ToInt32(_valores[9]), Convert.ToInt32(_valores[14]), mensaje);
                    break;
                case 3:
                     cambiarEstado(eesDomains.Default.ActualizarDatosCliente, buscarPalabras(NombrePlantilla), ticket, Convert.ToInt32(_valores[10]), Convert.ToInt32(_valores[11]), Convert.ToInt32(_valores[15]), mensaje);
                    break;
            }
        }

        public void cambiarEstado( int idproceso, string palabra,Ticket ticket,int estadoConforme,int estadoRechazo,int estadoMesaControl ,string mensaje)
        {
            try
            {
                if (palabra == "CONFORME")
                    _robot.SaveTicketNextState(ticket, estadoConforme);
                else if (palabra == "RECHAZO")
                    _robot.SaveTicketNextState(ticket, estadoRechazo);
                
                if(mensaje!="")
                    _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, mensaje), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == estadoMesaControl).Id);

            }
            catch (Exception ex)
            {
                _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, ex.Message), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == estadoMesaControl).Id);
            }

        }





    }
}
