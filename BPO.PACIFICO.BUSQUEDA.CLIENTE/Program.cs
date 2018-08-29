using everis.Ees.Proxy;
using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using Robot.Util.Nacar;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BPO.PACIFICO.BUSQUEDA.CLIENTE
{
    class Program : IRobot
    {
        private static BaseRobot<Program> _robot = null;
        private static IWebDriver _driverGlobal = new InternetExplorerDriver();
        string campo = string.Empty;
        string texto = string.Empty;
        private static Functions _Funciones;

        static string[] _valoresTickets = new string[10];
        static string[] _valoresTickets_Ident = new string[10];

        Ticket ticket = new Ticket();
        List<TicketValue> ticketVa = null;
        #region ParametrosRobot
        private string _urlContactManager = string.Empty;
        private string _usuarioContactManager = string.Empty;
        private string _contraseñaContactManager = string.Empty;
        private string _ToblaPersonaCampos = string.Empty;
        private string _ToblaPersonaPosicion = string.Empty;
        #endregion


        static void Main(string[] args)
        {

            _Funciones = new Functions();
            _robot = new BaseRobot<Program>(args);
            _robot.Start();
        }

        protected override void Start()
        {
            try
            {
                GetParameterRobots();
                ProcesarTicket();
            }
            catch (Exception ex)
            {
                LogFailProcess(Constants.MSG_ERROR_EVENT_PROCESS_KEY, ex);
            }

        }

        private void ProcesarTicket()
        {
            ticket = _robot.Tickets.FirstOrDefault();

            //= ticket.TicketValues[0].Value
            AbrirSelenium();
            NavegarUrl();
            Login();
            ValoresTicketsRobots();
            AccedientoContactManager();

        }

        private void ValoresTicketsRobots()
        {
            try
            {


                //Identificador  Tipo Contacto 
                _valoresTickets_Ident[0] = "Empresa";
                //Tipo Documento
                _valoresTickets_Ident[1] = "RUC";
                //iden
                _valoresTickets[1] = "20100049181";
              



                ticket = _robot.Tickets.FirstOrDefault();
                ticketVa = _robot.GetDataQueryTicketValue().Where(a => a.TicketId == ticket.Id).ToList();

            }
            catch (Exception ex)
            {
                throw new Exception("Error al Obtener los Valores del robot", ex);
            }
        }


        private void GetParameterRobots()
        {
            try
            {
                _urlContactManager = _robot.GetValueParamRobot("URLContactManager").ValueParam;
                _usuarioContactManager = _robot.GetValueParamRobot("UsuarioContactManager").ValueParam;
                _contraseñaContactManager = _robot.GetValueParamRobot("PasswordContactManager").ValueParam;
                _ToblaPersonaCampos = _robot.GetValueParamRobot("TablaPersonaCampo").ValueParam;
                _ToblaPersonaPosicion = _robot.GetValueParamRobot("TablaPersonaPosicion").ValueParam;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Obtener los parametros del robot", ex);
            }
        }

        private void AbrirSelenium()
        {
            //LogInfoStep(5);//id referencial msje Log "Iniciando la carga Internet Explorer"
            try
            {
                _Funciones.AbrirSelenium(ref _driverGlobal);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Iniciar Internet Explorer", ex);
            }
            //LogInfoStep(6);//id referencial msje Log "Finalizando la carga Internet Explorer"

        }

        private void NavegarUrl()
        {

            try
            {
                //LogInfoStep(5);//id referencial msje Log "Iniciando acceso al sitio policenter"
                _Funciones.NavegarUrlPolicyCenter(_driverGlobal, _urlContactManager);
            }
            catch (Exception ex)
            {
                throw new Exception("No se puede acceder al sitio policycenter", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Finalizando acceso al sitio policenter"


        }

        private void Login()
        {

            try
            {
                //LogInfoStep(5);//id referencial msje Log "Iniciando login policenter"

                _Funciones.LoginPolicyCenter(_driverGlobal, _usuarioContactManager, _contraseñaContactManager);
            }
            catch (Exception ex)
            {
                throw new Exception("No se puede acceder al sistema policycenter", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Finalizacion login policenter"


        }

        private void AccedientoContactManager()
        {
            if (_valoresTickets[0] != "")
            {
                _driverGlobal.FindElement(By.Name("ABContactSearch:ABContactSearchScreen:ContactSearchDV:PublicID")).SendKeys(_valoresTickets[0]);
            }
            else
            {

                IngresarTextoSelect("ABContactSearch:ABContactSearchScreen:ContactSearchDV:ContactSubtype", _valoresTickets_Ident[0]);
                IngresarTextoSelect("ABContactSearch:ABContactSearchScreen:ContactSearchDV:PrimaryOfficialIDTypeExt", _valoresTickets_Ident[1]);

                _driverGlobal.FindElement(By.Id("ABContactSearch:ABContactSearchScreen:ContactSearchDV:TaxID")).SendKeys(_valoresTickets[1]);
                _Funciones.Esperar(2);
            }

            _driverGlobal.FindElement(By.ClassName("bigButton_link")).Click();
            _Funciones.Esperar(2);

            _driverGlobal.FindElement(By.Id("ABContactSearch:ABContactSearchScreen:ContactSearchResultsLV:0:DisplayName")).Click();
            _Funciones.Esperar(2);


            _driverGlobal.FindElement(By.Id("ContactDetail:ABContactDetailScreen:ContactBasicsDV_tb:Edit")).Click();
            _Funciones.Esperar(3);

         
        }

        public void IngresarTextoSelect(String name, String texto)
        {
            SelectElement elemen = new SelectElement(_driverGlobal.FindElement(By.Name(name)));
            elemen.SelectByText(texto);
            _Funciones.Esperar(2);
        }






    }
}
