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

namespace BPO.PACIFICO.REHABILITAR
{
    class Program : IRobot
    {
        private static BaseRobot<Program> _robot = null;
        private static IWebDriver _driverGlobal = null;
        private static IWebElement element;
        private static Functions _Funciones;

        #region ParametrosRobot
        private string _urlPolicyCenter = string.Empty;
        private string _usuarioPolicyCenter = string.Empty;
        private string _contraseñaPolicyCenter = string.Empty;
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
            RehabilitarPoliza(ticket);
        }
        private void RehabilitarPoliza(Ticket ticket)
        {
            try
            {
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions")).Click();
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions:PolicyFileMenuActions_NewWorkOrder:PolicyFileMenuActions_ReinstatePolicy")).Click();
                _Funciones.Esperar(5);

                string _motivoIdElement = "ReinstatementWizard:ReinstatementWizard_ReinstatePolicyScreen:ReinstatePolicyDV:Reason";

                string _descripcionMotivo = "";

                int _idCampoDominioMotivo = Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == 1054).Value.ToString());


                string _textoDominioMotivo = _Funciones.ObtenerValorDominio(ticket, _idCampoDominioMotivo);
                _Funciones.SeleccionarCombo(_driverGlobal, _motivoIdElement, _textoDominioMotivo);
                _Funciones.Esperar(2);

                _driverGlobal.FindElement(By.Id("ReinstatementWizard:ReinstatementWizard_ReinstatePolicyScreen:ReinstatePolicyDV:ReasonDescription")).SendKeys(string.Concat(_descripcionMotivo));
                _Funciones.Esperar(2);
                _driverGlobal.FindElement(By.Id("ReinstatementWizard:ReinstatementWizard_ReinstatePolicyScreen:JobWizardToolbarButtonSet:QuoteOrReview")).Click();
                _Funciones.Esperar(1);
                _driverGlobal.FindElement(By.Id("ReinstatementWizard:ReinstatementWizard_QuoteScreen:JobWizardToolbarButtonSet:Reinstate")).Click();
                _Funciones.Esperar(1);
                _driverGlobal.SwitchTo().Alert().Accept();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Rehabilitar la poliza en el sistema policycenter", ex);
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
                _Funciones.NavegarUrlPolicyCenter(_driverGlobal, _urlPolicyCenter);
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
                _Funciones.LoginPolicyCenter(_driverGlobal, _usuarioPolicyCenter, _contraseñaPolicyCenter);
            }
            catch (Exception ex)
            {
                throw new Exception("No se puede acceder al sistema policycenter", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Finalizacion login policenter"


        }
        private void BuscarPoliza(Ticket ticket)
        {
            _numeroPoliza = ticket.TicketValues.FirstOrDefault(np => np.FieldId == 5).Value.ToString();


            try
            {
                //LogInfoStep(5);//id referencial msje Log "Iniciando busqueda de poliza"
                if (!string.IsNullOrEmpty(_numeroPoliza))
                {
                    _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, _numeroPoliza);
                }
                //LogInfoStep(5);//id referencial msje Log "Finalizando busqueda de poliza"
            }
            catch (Exception ex)
            {
                throw new Exception("Error al buscar el numero de poliza policycenter", ex);
            }


        }
        private void GetParameterRobots()
        {
            try
            {
                _urlPolicyCenter = _robot.GetValueParamRobot("URLPolyCenter").ValueParam;
                _usuarioPolicyCenter = _robot.GetValueParamRobot("UsuarioPolyCenter").ValueParam;
                _contraseñaPolicyCenter = _robot.GetValueParamRobot("PasswordPolyCenter").ValueParam;

            }
            catch (Exception ex)
            {
                throw new Exception("Error al Obtener los parametros del robot", ex);
            }
        }
    }
}
