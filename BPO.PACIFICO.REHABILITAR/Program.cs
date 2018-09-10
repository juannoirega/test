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
        private int _estadoError;
        private int _estadoFinal;
        #endregion
        #region VariablesGLoables
        private int _reprocesoContador = 0;
        private int _idEstadoRetorno = 0;
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
                LogFailStep(30, ex);
            }

            foreach (Ticket ticket in _robot.Tickets)
            {
                try
                {
                    var valoresReprocesamiento = _Funciones.ObtenerValoresReprocesamiento(ticket);
                    if (valoresReprocesamiento.Count > 0) { _reprocesoContador = valoresReprocesamiento[0]; _idEstadoRetorno = valoresReprocesamiento[1]; }
                    ProcesarTicket(ticket);
                }
                catch (Exception ex)
                {
                    LogFailStep(30, ex);
                    _reprocesoContador++;
                    _Funciones.GuardarValoresReprocesamiento(ticket, _reprocesoContador, _idEstadoRetorno);
                    _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, ex.Message), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == _estadoError).Id);
                }
                finally
                {
                    _Funciones.CerrarDriver(_driverGlobal);
                }
            }
        }

        private void ProcesarTicket(Ticket ticket)
        {
            _Funciones.AbrirSelenium(ref _driverGlobal);
            _Funciones.NavegarUrlPolicyCenter(_driverGlobal, _urlPolicyCenter);
            _Funciones.LoginPolicyCenter(_driverGlobal, _usuarioPolicyCenter, _contraseñaPolicyCenter);
            _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, ticket.TicketValues.FirstOrDefault(tv => tv.FieldId == eesFields.Default.poliza_nro).Value);
            RehabilitarPoliza(ticket);
            if (_reprocesoContador > 0)
            {
                _reprocesoContador = 0;
                _idEstadoRetorno = 0;
                _Funciones.GuardarValoresReprocesamiento(ticket, _reprocesoContador, _idEstadoRetorno);
            }
            _robot.SaveTicketNextState(ticket, _estadoFinal);
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

        private void GetParameterRobots()
        {
            _urlPolicyCenter = _robot.GetValueParamRobot("URLPolyCenter").ValueParam;
            _usuarioPolicyCenter = _robot.GetValueParamRobot("UsuarioPolyCenter").ValueParam;
            _contraseñaPolicyCenter = _robot.GetValueParamRobot("PasswordPolyCenter").ValueParam;
            _estadoError = Convert.ToInt32(_robot.GetValueParamRobot("EstadoError").ValueParam);
            _estadoFinal = Convert.ToInt32(_robot.GetValueParamRobot("EstadoSiguiente").ValueParam);
            LogEndStep(4);
        }
    }
}
