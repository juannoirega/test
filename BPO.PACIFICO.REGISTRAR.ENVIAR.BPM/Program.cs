using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using Robot.Util.Nacar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPO.PACIFICO.REGISTRAR.ENVIAR.BPM
{
    public class Program : IRobot
    {
        #region "PARÁMETROS"
        private static BaseRobot<Program> _robot = null;
        private static IWebDriver _oDriver = null;
        private static IWebElement _oElement;
        private static int _nIdEstadoError;
        private static int _nIdEstadoFinal;
        private static string _cUrlOnBase = string.Empty;
        private static string _cUsuarioOnBase = string.Empty;
        private static string _cContraseñaOnBase = string.Empty;
        private static StateAction _oMesaControl;
        private static StateAction _oRegistro;
        private static List<StateAction> _oAcciones;
        #endregion

        static void Main(string[] args)
        {
            _robot = new BaseRobot<Program>(args);
            _oDriver = new FirefoxDriver();
            //_oDriver = new ChromeDriver();
            _robot.Start();
        }

        protected override void Start()
        {
            if (_robot.Tickets.Count < 1)
                return;

            ObtenerParametros();
            LogStartStep(4);
            foreach (Ticket oTicket in _robot.Tickets)
            {
                try
                {
                    _oAcciones = _robot.GetNextStateAction(oTicket);
                    _oMesaControl = _oAcciones.Where(a => a.ActionId == _nIdEstadoError).SingleOrDefault();
                    _oRegistro = _oAcciones.Where(b => b.ActionId == _nIdEstadoFinal).SingleOrDefault();
                    ProcesarTicket(oTicket);
                }
                catch (Exception Ex)
                {
                    LogFailStep(12, Ex);
                }
            }
        }

        //Obtiene valores para los parámetros del Robot desde EES:
        private void ObtenerParametros()
        {
            try
            {
                //Parámetros del Robot Procesamiento de Datos:
                _nIdEstadoError = Convert.ToInt32(_robot.GetValueParamRobot("EstadoError").ValueParam);
                _nIdEstadoFinal = Convert.ToInt32(_robot.GetValueParamRobot("EstadoFinal").ValueParam);
                _cUrlOnBase = _robot.GetValueParamRobot("URLOnBase").ValueParam;
                _cUsuarioOnBase = _robot.GetValueParamRobot("UsuarioOnBase").ValueParam;
                _cContraseñaOnBase = _robot.GetValueParamRobot("ContraseñaOnBase").ValueParam;
            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
            }
        }

        //Inicia el procesamiento de tickets:
        private void ProcesarTicket(Ticket oTicketDatos)
        {
            try
            {
                //Valida campos no vacíos:
                if (!ValidarVacios(oTicketDatos))
                {
                    IngresarBPM();
                    RegistrarBPM();
                }
                else
                {
                    //Enviar a mesa de control:
                    CambiarEstadoTicket(oTicketDatos, _oMesaControl);
                }
            }
            catch (Exception Ex)
            {
                LogFailStep(17, Ex);
            }
        }

        //Valida que no tenga campos vacíos:
        private bool ValidarVacios(Ticket oTicketDatos)
        {
            int[] oCampos = new int[] { eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado,
                                        eesFields.Default.tipo_de_linea};
            return true;
        }

        //Envía el ticket al siguiente estado:
        private void CambiarEstadoTicket(Ticket oTicket, StateAction oAccion)
        {
            //Estado = 1: Mesa de Control, Estado = 2: Notificación de Correo.
            _robot.SaveTicketNextState(oTicket, oAccion.Id);
        }

        //Ingresa al sistema OnBase:
        private void IngresarBPM()
        {
            try
            {
                _oDriver.Url = _cUrlOnBase;
                _oDriver.Manage().Window.Maximize();
                _oDriver.SwitchTo().Frame(_oDriver.FindElement(By.Id("NavPanelIFrame")));
                _oDriver.FindElement(By.LinkText("VSER_Consulta de Solicitudes 1"));

                _oElement.Click();
            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
            }
        }

        //Realiza el registro de anulación en OnBase:
        private void RegistrarBPM()
        {
            try
            {

            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
            }
        }
    }
}
