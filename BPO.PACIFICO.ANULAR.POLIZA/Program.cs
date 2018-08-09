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
    class Program: IRobot
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
                   // ProcesarTicket(ticket);
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
        private void AbrirSelenium()
        {
            //LogInfoStep(5);//id referencial msje Log "Iniciando la carga Internet Explorer"
            try
            {
                Functions.AbrirSelenium(ref _driverGlobal);
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
                Functions.NavegarUrl(_driverGlobal, _url);
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
                Functions.Login(_driverGlobal, _usuario, _contraseña);
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
                _numeroPoliza = ticket.TicketValues.Where(np => np.FieldId == 5).ToString();

                Functions.BuscarPoliza(_driverGlobal, _numeroPoliza);
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
