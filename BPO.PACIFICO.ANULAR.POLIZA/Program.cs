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
        private static Functions _Funciones;
        #region ParametrosRobot
        private string _urlPolicyCenter = string.Empty;
        private string _usuarioPolicyCenter = string.Empty;
        private string _contraseñaPolicyCenter = string.Empty;
        private string _usuarioBcp = string.Empty;
        private string _contraseñaBcp = string.Empty;
        private string _urlBcp = string.Empty;
        private int _estadoError;
        private int _estadoFinal;
        #endregion
        #region VariablesGLoables
        private string _numeroPoliza = string.Empty;
        //variable temporal
        //private static bool _esPortalBcp = false;

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
                    ProcesarTicket(ticket);
                }
                catch (Exception ex)
                {
                    LogFailStep(30, ex);
                    _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, ex.Message), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == _estadoError).Id);
                    //Enviar a mesa control con mmensaje
                    //capturar imagen
                }
                finally
                {
                    if (_driverGlobal != null)
                        _driverGlobal.Quit();
                }
            }
        }
        private void ProcesarTicket(Ticket ticket)
        {
            //falta verificar cual sera el id del campo que confirmara si es portalbc o no**reemplazar por el "1"
            //_esPortalBcp = ticket.TicketValues.FirstOrDefault(tv => tv.FieldId == 1).Value.ToString() == "True" ? true : false;
            AbrirSelenium();
            NavegarUrl();
            Login();
            BuscarPoliza(ticket);
            AnularPoliza(ticket);
        }
        private void AnularPoliza(Ticket ticket)
        {
            //if (_esPortalBcp)
            //    AnularPolizaPortalBcp();
            //else
                AnularPolizaPolicyCenter(ticket);
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
            //if (_esPortalBcp)
            //{
            //    try
            //    {
            //        //LogInfoStep(5);//id referencial msje Log "Iniciando acceso al sitio policenter"
            //        _Funciones.NavegarUrlPortalBcp(_driverGlobal, _urlBcp);
            //    }
            //    catch (Exception ex)
            //    {
            //        throw new Exception("No se puede acceder al sitio portal bcp", ex);
            //    }
            //    //LogInfoStep(5);//id referencial msje Log "Finalizando acceso al sitio policenter"
            //}
            //else
            //{
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
            //}

        }

        private void Login()
        {
            //if (_esPortalBcp)
            //{
            //    try
            //    {
            //        //LogInfoStep(5);//id referencial msje Log "Iniciando login policenter"
            //        _Funciones.LoginPortalBcp(_driverGlobal, _usuarioBcp, _contraseñaBcp);
            //    }
            //    catch (Exception ex)
            //    {
            //        throw new Exception("No se puede acceder al sistema portal bcp", ex);
            //    }
            //    //LogInfoStep(5);//id referencial msje Log "Finalizacion login policenter"
            //}
            //else
            //{
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
            //}

        }
        private void BuscarPoliza(Ticket ticket)
        {
            _numeroPoliza = ticket.TicketValues.FirstOrDefault(np => np.FieldId == 5).Value.ToString();

            //if (_esPortalBcp)
            //{
            //    try
            //    {
            //        //LogInfoStep(5);//id referencial msje Log "Iniciando busqueda de poliza"
            //        if (!string.IsNullOrEmpty(_numeroPoliza))
            //        {
            //            _Funciones.BuscarPolizaPortalBcp(_driverGlobal, _numeroPoliza);
            //        }
            //        //LogInfoStep(5);//id referencial msje Log "Finalizando busqueda de poliza"
            //    }
            //    catch (Exception ex)
            //    {
            //        throw new Exception("Error al buscar el numero de poliza portal bcp", ex);
            //    }
            //}
            //else
            //{
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
            //}

        }

        private void AnularPolizaPortalBcp()
        {
            try
            {
                _driverGlobal.FindElement(By.XPath("//img[contains(@id,'ctl00_ContentPlaceHolder1_gvPolizas_ctl03_imgVer')]")).Click();
                _Funciones.Esperar(2);
                _driverGlobal.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlTipoMod']/option[2]")).Click();
                _Funciones.Esperar(3);
                _driverGlobal.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlMotivo_03']/option[2]")).Click();
                //Falta confirmar la anulacion de poliza portal bcp
            }
            catch (Exception ex)
            {
                throw new Exception("Error al anular la poliza en el sistema portal bcp", ex);
            }

        }
        private void AnularPolizaPolicyCenter(Ticket ticket)
        {
            try
            {
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions")).Click();
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions:PolicyFileMenuActions_NewWorkOrder:PolicyFileMenuActions_CancelPolicy")).Click();
                _Funciones.Esperar(5);

                string _solicitanteIdElement = "StartCancellation:StartCancellationScreen:CancelPolicyDV:Source";
                string _motivoIdElement = "StartCancellation:StartCancellationScreen:CancelPolicyDV:Reason2";
                string _reembolsoIdElement = "StartCancellation:StartCancellationScreen:CancelPolicyDV:CalcMethod";
                string _descripcionMotivo = "SE DEJA CONSTANCIA POR EL PRESENTE ENDOSO QUE, LA POLIZA DEL RUBRO QUEDA CANCELADA, NULA Y SIN VALOR PARA TODOS SUS EFECTOS A PARTIR DEL";


                int _idCampoDominioSolicitante = Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == 1051).Value.ToString());
                int _idCampoDominioMotivo = Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == 11).Value.ToString());
                int _idCampoDominioReembolso = Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == 16).Value.ToString());

                string _textoDominioSolicitante = _Funciones.ObtenerValorDominio(ticket, _idCampoDominioSolicitante);
                _Funciones.SeleccionarCombo(_driverGlobal, _solicitanteIdElement, _textoDominioSolicitante);
                _Funciones.Esperar(2);

                string _textoDominioMotivo = _Funciones.ObtenerValorDominio(ticket, _idCampoDominioMotivo);
                _Funciones.SeleccionarCombo(_driverGlobal, _motivoIdElement, _textoDominioMotivo);
                _Funciones.Esperar(2);

                string _fechaEfectivaCancelacion = _Funciones.ObtenerValorElemento(_driverGlobal, "StartCancellation:StartCancellationScreen:CancelPolicyDV:CancelDate_date");

                _driverGlobal.FindElement(By.Id("StartCancellation:StartCancellationScreen:CancelPolicyDV:ReasonDescription")).SendKeys(string.Concat(_descripcionMotivo, " ", _fechaEfectivaCancelacion));
                _Funciones.Esperar(2);

                string _textoDominioReembolso = _Funciones.ObtenerValorDominio(ticket, _idCampoDominioReembolso);
                _Funciones.SeleccionarCombo(_driverGlobal, _reembolsoIdElement, _textoDominioReembolso);
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
                throw new Exception("Error al Anular la poliza en el sistema policycenter", ex);
            }

        }
        private void GetParameterRobots()
        {
            try
            {
                _urlPolicyCenter = _robot.GetValueParamRobot("URLPolyCenter").ValueParam;
                _usuarioPolicyCenter = _robot.GetValueParamRobot("UsuarioPolyCenter").ValueParam;
                _contraseñaPolicyCenter = _robot.GetValueParamRobot("PasswordPolyCenter").ValueParam;
                //_usuarioBcp = _robot.GetValueParamRobot("UsuarioBcp").ValueParam;
                //_contraseñaBcp = _robot.GetValueParamRobot("PasswordBcp").ValueParam;
                //_urlBcp = _robot.GetValueParamRobot("URLBcp").ValueParam;
                _estadoError = Convert.ToInt32(_robot.GetValueParamRobot("EstadoErrorAP").ValueParam);
                _estadoFinal = Convert.ToInt32(_robot.GetValueParamRobot("EstadoSiguienteAP").ValueParam);
                LogEndStep(4);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Obtener los parametros del robot", ex);
            }
        }
    }
}
