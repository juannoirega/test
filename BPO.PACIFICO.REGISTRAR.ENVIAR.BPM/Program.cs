using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
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
        private static string _cTipoSolicitud1 = string.Empty;
        private static string _cTipoSolicitud2 = string.Empty;
        private static StateAction _oMesaControl;
        private static StateAction _oRegistro;
        private static List<StateAction> _oAcciones;
        private static Functions _Funciones;
        private string cContratante;
        #endregion

        static void Main(string[] args)
        {
            _robot = new BaseRobot<Program>(args);
            _oDriver = new FirefoxDriver();
            _Funciones = new Functions();
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
                _cTipoSolicitud1 = _robot.GetValueParamRobot("TipoSolicitud1").ValueParam;
                _cTipoSolicitud2 = _robot.GetValueParamRobot("TipoSolicitud2").ValueParam;
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
                    _Funciones.IngresarBPM(_cUrlOnBase, _cUsuarioOnBase, _cContraseñaOnBase);
                    RegistrarBPM(oTicketDatos);
                }
                else
                {
                    //Enviar a mesa de control:
                    CambiarEstadoTicket(oTicketDatos, _oMesaControl);
                }
            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
            }
        }

        //Valida que no tenga campos vacíos:
        private bool ValidarVacios(Ticket oTicketDatos)
        {
            int[] oCampos = new int[] { eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado};
            return true;
        }

        //Envía el ticket al siguiente estado:
        private void CambiarEstadoTicket(Ticket oTicket, StateAction oAccion)
        {
            //Estado = 1: Mesa de Control, Estado = 2: Notificación de Correo.
            _robot.SaveTicketNextState(oTicket, oAccion.Id);
        }

        //Realiza el registro de anulación en OnBase:
        private void RegistrarBPM(Ticket oTicketDatos)
        {
            try
            {
                _oDriver = new FirefoxDriver();
                //_oDriver.SwitchTo().Frame(_oDriver.FindElement(By.LinkText("lbCustomQueries")));
                //Clic en pestaña Consultas Personalizadas:
                _oDriver.SwitchTo().Frame(_oDriver.FindElement(By.Id("subDownArrow")));

                //_oDriver.SwitchTo().Frame(_oDriver.FindElement(By.LinkText("VSER_Formulario Solicitud de Polizas")));
                //Click en "Nuevo Formulario":
                _oDriver.FindElement(By.XPath("//*[@id='SubMenuOptionsTable']/tbody/tr[2]/td"));

                //Clic en VSER_Formulario Solicitud de Polizas:
                _oDriver.FindElement(By.ClassName("formName"));

                //Selecciona tipo de solicitud:
                SelectElement _oTipoSolicitud = new SelectElement(_oDriver.FindElement(By.ClassName("ui-corner-all")));
                _oTipoSolicitud.SelectByText(_cTipoSolicitud2);

                //--- otra forma:
                _oDriver.FindElement(By.Id("tipodesolicitud_input")).Click();
                _oDriver.FindElement(By.Id("tipodesolicitud_input")).SendKeys(_cTipoSolicitud2); //EN DURO SE ESTÁ ENVIANDO "ENDOSO".
                Repeticiones(1, "tipodesolicitud_input", Keys.Down);
                _oDriver.FindElement(By.Id("tipodesolicitud_input")).SendKeys(Keys.Enter);                

                //Ingresa fecha hora de email:
                _oDriver.FindElement(By.Id("fechayhoraderecepción_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_hora_de_email).Value);

                //Selecciona Línea de Negocio:
                //necesario validar por el tipo de línea:
                SelectElement _oLineaNegocio = new SelectElement(_oDriver.FindElement(By.ClassName("ui-corner-all")));
                _oLineaNegocio.SelectByText("LINEAS PERSONALES");

                //---> otra forma:
                _oDriver.FindElement(By.Id("linea_negocio_input")).Click();
                Repeticiones(2, "linea_negocio_input", Keys.Down);
                _oDriver.FindElement(By.Id("linea_negocio_input")).SendKeys(Keys.Enter);

                //Selecciona Producto:
                SelectElement _oProducto = new SelectElement(_oDriver.FindElement(By.ClassName("ui-corner-all")));
                _oProducto.SelectByText("ACCIDENTES INDIVIDUALES");

                //---> otra forma.
                _oDriver.FindElement(By.Id("producto_input")).Click();
                Repeticiones(11, "producto_input", Keys.Down);
                _oDriver.FindElement(By.Id("producto_input")).SendKeys(Keys.Enter);
                _Funciones.Esperar(3);

                //Seleccionar Tipo de endoso: Necesario validar por el endoso en curso
                //_oDriver.FindElement(By.Id("ui-active-menuitem")).Click();
                SelectElement _oTipoEndoso = new SelectElement(_oDriver.FindElement(By.ClassName("ui-corner-all")));
                _oTipoEndoso.SelectByText("ANULACIÓN DE PÓLIZA");
                _Funciones.Esperar();

                //Seleccionar Motivo de Anulación:
                string cMotivo = oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == 1).Value; //falta ID del motivo.
                SelectElement _oMotivoAnulacion = new SelectElement(_oDriver.FindElement(By.ClassName("ui-corner-all")));
                _oMotivoAnulacion.SelectByText("ANULACIÓN MASIVA POR TRAMA");

                //Ingresar número de asegurados:
                _oDriver.FindElement(By.Id("nrounidnroasegnrocertinrobienes_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == 2).Value); //falta ID de nro asegurados/vehiculos/bienes

                string cContratante = oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre_contratante).Value;
                string cAsegurado = oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre_asegurado).Value;
                     
            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
            }
        }

        private void Repeticiones(int nVeces, string cElemento, string cKeys)
        {
            for (int i = 0; i < nVeces; i++)
            {
                _oDriver.FindElement(By.Id(cElemento)).SendKeys(cKeys);
            }
        }
    }
}
