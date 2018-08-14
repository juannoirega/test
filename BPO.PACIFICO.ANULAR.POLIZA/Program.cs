using everis.Ees.Proxy;
using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using Robot.Util.Nacar;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BPO.PACIFICO.ANULAR.POLIZA
{
    class Program : IRobot
    {
        private static BaseRobot<Program> _robot = null;
        private static IWebDriver _driverGlobal = null;
        private static IWebElement element;
        private static Functions _Funciones;
        #region ParametrosRobot
        private string _url = string.Empty;
        private string _usuario = string.Empty;
        private string _contraseña = string.Empty;
        #endregion
        #region VariablesGLoables
        private string _numeroPoliza = string.Empty;
        private string _producto = string.Empty;
        private string _inicioVigencia = string.Empty;
        private string _finVigencia = string.Empty;
        private string _numeroAgente = string.Empty;
        private string _agente = string.Empty;
        private string _tipo = string.Empty;
        private string _estado = string.Empty;
        private string _tipoVigencia = string.Empty;
        private string _numeroCanal = string.Empty;
        private string _nombreContratante = string.Empty;
        private string _nombreAsegurado = string.Empty;
        //private static string _numeroDniContratante = string.Empty;
        //Pendiente
        private string _numeroVehiculos = string.Empty;
        private string _numeroAsegurados = string.Empty;

        #endregion

        static void Main(string[] args)
        {
            _Funciones = new Functions();
            _robot = new BaseRobot<Program>(args);
            _robot.Start();
        }

        protected override void Start()
        {
            if (_robot.Tickets.Count < 1)
                return;

            LogStartStep(2);
            try
            {
                GetParameterRobots();
            }
            catch (Exception ex)
            {
                LogFailProcess(Constants.MSG_ERROR_EVENT_PROCESS_KEY, ex);
            }

            foreach (Ticket ticket in _robot.Tickets)
            {
                try
                {
                    ProcesarTicket(ticket);
                }
                catch (Exception ex)
                {
                    LogFailProcess(Constants.MSG_ERROR_EVENT_PROCESS_KEY, ex);
                    //Enviar a mesa control con mmensaje
                    //capturar imagen
                }
                finally
                {
                    if (_driverGlobal != null)
                        _driverGlobal.Quit();

                    LogEndStep(Constants.MSG_PROCESS_ENDED_KEY);
                }
            }
        }

        private void ProcesarTicket(Ticket ticket)
        {
            AbrirSelenium();
            NavegarUrl();
            Login();
            BuscarPoliza(ticket);
            AnularPoliza(ticket);
        }
        private void AnularPoliza(Ticket ticket)
        {
            try
            {
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions")).Click();
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions:PolicyFileMenuActions_NewWorkOrder:PolicyFileMenuActions_CancelPolicy")).Click();
                _Funciones.Esperar();

                string _solicitanteIdElement = "StartCancellation:StartCancellationScreen:CancelPolicyDV:Source";
                string _motivoIdElement = "StartCancellation:StartCancellationScreen:CancelPolicyDV:Reason2";
                string _reembolsoIdElement = "StartCancellation:StartCancellationScreen:CancelPolicyDV:CalcMethod";
                string _descripcionMotivo = "SE DEJA CONSTANCIA POR EL PRESENTE ENDOSO QUE, LA POLIZA DEL RUBRO QUEDA CANCELADA, NULA Y SIN VALOR PARA TODOS SUS EFECTOS A PARTIR DEL";

                string _solicitante = "Compañía de seguros";
                _Funciones.SeleccionarCombo(_driverGlobal, _solicitanteIdElement, _solicitante);
                _Funciones.Esperar(2);

                string _motivo = "CANCELÓ CRÉDITO";
                _Funciones.SeleccionarCombo(_driverGlobal, _motivoIdElement, _motivo);
                _Funciones.Esperar(2);

                string _fechaEfectivaCancelacion = _Funciones.ObtenerValorElemento(_driverGlobal, "StartCancellation:StartCancellationScreen:CancelPolicyDV:CancelDate_date");

                _driverGlobal.FindElement(By.Id("StartCancellation:StartCancellationScreen:CancelPolicyDV:ReasonDescription")).SendKeys(string.Concat(_descripcionMotivo, " ", _fechaEfectivaCancelacion));
                _Funciones.Esperar(2);

                string _reembolso = "Devolución 100%";
                _Funciones.SeleccionarCombo(_driverGlobal, _reembolsoIdElement, _reembolso);
                _Funciones.Esperar(2);

                _driverGlobal.FindElement(By.Id("StartCancellation:StartCancellationScreen:NewCancellation")).Click();
                _Funciones.Esperar(1);

                _driverGlobal.FindElement(By.Id("CancellationWizard:CancellationWizard_QuoteScreen:JobWizardToolbarButtonSet:BindOptions_arrow")).Click();
                _driverGlobal.FindElement(By.Id("CancellationWizard:CancellationWizard_QuoteScreen:JobWizardToolbarButtonSet:BindOptions:CancelNow")).Click();
                _Funciones.Esperar(1);

                _driverGlobal.SwitchTo().Alert().Accept();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Anular la Poliza", ex);
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
            //LogInfoStep(5);//id referencial msje Log "Iniciando acceso al sitio policenter"

            try
            {
                _Funciones.NavegarUrl(_driverGlobal, _url);
            }
            catch (Exception ex)
            {
                throw new Exception("No se pudo acceder al sitio policenter", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Finalizando acceso al sitio policenter"


        }
        private void Login()
        {
            //LogInfoStep(5);//id referencial msje Log "Iniciando login policenter"

            try
            {
                _Funciones.Login(_driverGlobal, _usuario, _contraseña);
            }
            catch (Exception ex)
            {
                throw new Exception("No se pudo acceder al sistema policenter", ex);

            }
            //LogInfoStep(5);//id referencial msje Log "Finalizacion login policenter"

        }
        private void BuscarPoliza(Ticket ticket)
        {
            //LogInfoStep(5);//id referencial msje Log "Iniciando busqueda de poliza"

            try
            {
                //obtener el numero  de Poliza
                _numeroPoliza = ticket.TicketValues.FirstOrDefault(np => np.FieldId == 5).Value.ToString();

                _Funciones.BuscarPoliza(_driverGlobal, _numeroPoliza);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al buscar el numero de poliza", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Finalizando busqueda de poliza"

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
