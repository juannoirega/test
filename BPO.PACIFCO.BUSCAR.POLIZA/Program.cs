using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BPO.PACIFCO.BUSCAR.POLIZA
{

    class Program : IRobot
    {
        private static BaseRobot<Program> _robot = null;
        private static IWebDriver _driverGlobal = null;
        private static IWebElement element;




        #region ParametrosRobot
        private string _url = string.Empty;
        private string _usuario = string.Empty;
        private string _contraseña = string.Empty;
        #endregion
        #region VariablesGLoables
        private static string _producto = string.Empty;
        private static string _inicioVigencia = string.Empty;
        private static string _finVigencia = string.Empty;
        private static string _numeroAgente = string.Empty;
        private static string _agente = string.Empty;
        private static string _tipo = string.Empty;
        private static string _estado = string.Empty;
        private static string _tipoVigencia = string.Empty;
        private static string _numeroCanal = string.Empty;

        #endregion
        static void Main(string[] args)
        {
            //_robot = new BaseRobot<Program>(args);
            //_robot.Start();
            Login();
        }

        protected override void Start()
        {
            if (_robot.Tickets.Count < 1)
                return;

            LogStartProcess(1, 2);

            try
            {
                GetParameterRobots();
            }
            catch (Exception ex)
            {
                LogFailProcess(1, ex);
            }
            foreach (Ticket ticket in _robot.Tickets)
            {
                try
                {
                    ProcesarTicket(ticket);
                }
                catch (Exception ex)
                {
                    LogFailProcess(1, ex);
                }
            }
        }

        private void ProcesarTicket(Ticket ticket)
        {
            Login();
            BuscarPoliza();
            ObtenerDatos();
            GrabarInformacion(ticket);
        }

        static void Login()
        {
            try
            {
                //parametros en duro
                string url = ConfigurationManager.AppSettings["Url"];
                string usuario = ConfigurationManager.AppSettings["Usuario"];
                string clave = ConfigurationManager.AppSettings["Contraseña"];

                _driverGlobal = new InternetExplorerDriver();
                _driverGlobal.Url = url;
                _driverGlobal.Navigate().GoToUrl("javascript:document.getElementById('overridelink').click()");

                _driverGlobal.SwitchTo().DefaultContent();
                _driverGlobal.SwitchTo().Frame(_driverGlobal.FindElement(By.Id("top_frame")));

                element = _driverGlobal.FindElement(By.Id("Login:LoginScreen:LoginDV:username"));
                element.Clear();
                element.SendKeys(usuario);

                element = _driverGlobal.FindElement(By.Id("Login:LoginScreen:LoginDV:password"));
                element.Clear();
                element.SendKeys(clave);

                _driverGlobal.FindElement(By.Id("Login:LoginScreen:LoginDV:submit")).SendKeys(Keys.Enter);
                Thread.Sleep(300);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al intentar acceder al sistema", ex);
            }

        }
        static void BuscarPoliza()
        {
            try
            {
                //Desplegar menu Poliza
                _driverGlobal.FindElement(By.Id("TabBar:PolicyTab_arrow")).Click();
                //Seleccionar Input Poliza
                element = _driverGlobal.FindElement(By.Id("TabBar:PolicyTab:PolicyTab_PolicyRetrievalItem"));
                element.Clear();
                //deberia ser del ticket
                element.SendKeys("2500029553");
                //Buscar Poliza
                _driverGlobal.FindElement(By.Id("TabBar:PolicyTab:PolicyTab_PolicyRetrievalItem")).SendKeys(Keys.Enter);

                //Por ahora se llama aqui al metodo para las pruebas
                ObtenerDatos();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al buscar poliza", ex);
            }
        }

        static void ObtenerDatos()
        {
            try
            {
                //producto
                element = _driverGlobal.FindElement(By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_PolicyDV:Product"));
                _producto = element.GetAttribute("value");
                //Inicio Vigencia
                element = _driverGlobal.FindElement(By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:PolicyPerEffDate_date"));
                _inicioVigencia = element.GetAttribute("value");
                //Fin Vigencia
                element = _driverGlobal.FindElement(By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:PolicyPerExpirDate_date"));
                _finVigencia = element.GetAttribute("value");
                //Tipo
                element = _driverGlobal.FindElement(By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AssocJobDV:Type"));
                _tipo = element.GetAttribute("value");
                //Estado
                element = _driverGlobal.FindElement(By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AssocJobDV:state"));
                _estado = element.GetAttribute("value");
                //Tipo Vigencia
                element = _driverGlobal.FindElement(By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:validityType"));
                _tipoVigencia = element.GetAttribute("value");
                //Numero Canal y Nombre
                element = _driverGlobal.FindElement(By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_ProducerDV:PolicyInfoProducerInfoSummaryInputSet:SecondaryProducerCode"));
                string _canalCadenaCompleta = element.GetAttribute("value");
                string[] _arrayCanal = _canalCadenaCompleta.Split(' ');

                int c = 0;
                foreach (string item in _arrayCanal)
                {
                    if (c == 0) { _numeroCanal = item; c++; }
                }

                //Agente Numero y Nombre
                element = _driverGlobal.FindElement(By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_ProducerDV:PolicyInfoProducerInfoSummaryInputSet:ProducerCodeOfRecord"));
                string _agenteCadenaCompleta = element.GetAttribute("value");
                string[] _arrayAgente = _agenteCadenaCompleta.Split(' ');
                int i = 0;
                foreach (string item in _arrayAgente)
                {
                    if (i == 0) { _numeroAgente = item; i++; }
                    else { _agente = string.Concat(_agente, item, " "); }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Obtener los datos del sistema", ex);
            }
        }

        private void GrabarInformacion(Ticket ticket)
        {
            try
            {
                ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, FieldId = 1034, Value = _producto });
                ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, FieldId = 9, Value = _inicioVigencia });
                ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, FieldId = 8, Value = _finVigencia });
                ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, FieldId = 2, Value = _agente });
                ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, FieldId = 1036, Value = _numeroAgente });
                ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, FieldId = 26, Value = _tipo });
                ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, FieldId = 1033, Value = _tipoVigencia });
                ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, FieldId = 1032, Value = _estado });
            }
            catch (Exception ex)
            {
                throw new Exception("Ocurrio un Error al grabar la informacion en el ticket", ex);
            }
        }
        private void GetParameterRobots()
        {
            try
            {
                _url = _robot.GetValueParamRobot("URLPolyCenter").ValueParam;
                _usuario = _robot.GetValueParamRobot("UsuarioPolyCenter").ValueParam;
                _contraseña = _robot.GetValueParamRobot("PasswordPolyCenter").ValueParam;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Obtener los parametros del robot", ex);
            }
        }
    }
}
