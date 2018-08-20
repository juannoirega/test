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
        private static int _nIdEstadoError;
        private static int _nIdEstadoFinal;
        private static string _cUrlOnBase = string.Empty;
        private static string _cTipoSolicitud = string.Empty;
        private static int _nIntentosOnBase;
        private static string _cOpcionOnBase;
        private static string _cLineaPorDefecto;
        private static string _cRutaGeckodriver = string.Empty;
        private static StateAction _oMesaControl;
        private static StateAction _oRegistro;
        private static List<StateAction> _oAcciones;
        private static Functions _Funciones;
        private static string[] Usuarios;
        #endregion

        static void Main(string[] args)
        {
            try
            {
                _robot = new BaseRobot<Program>(args);
                _Funciones = new Functions();
                _robot.Start();
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex.Message);
            }
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
                    _Funciones.CerrarDriver(_oDriver);
                    return;
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
                _cTipoSolicitud = _robot.GetValueParamRobot("TipoSolicitud").ValueParam;
                _nIntentosOnBase = Convert.ToInt32(_robot.GetValueParamRobot("AccesoOnBase_Intentos").ValueParam);
                _cOpcionOnBase = _robot.GetValueParamRobot("OnBase_Opcion").ValueParam;
                _cLineaPorDefecto = _robot.GetValueParamRobot("LineaPorDefecto").ValueParam;
                _cRutaGeckodriver = _robot.GetValueParamRobot("Ruta_geckodriver").ValueParam;
            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
            }
        }

        //Obtiene los usuarios OnBase según Línea de negocio:
        private string[] UsuariosOnBase(Ticket oTicketDatos, int nIndice = 1)
        {
            if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_linea).Value == _cLineaPorDefecto)
            {
                //Usuario y contraseña de Autos:
                Usuarios = _robot.GetValueParamRobot("AccesoOnBase_Autos_" + nIndice).ValueParam.Split(',');
            }
            else
            {
                //Usuario y contraseña de Líneas Personales:
                Usuarios = _robot.GetValueParamRobot("AccesoOnBase_LLPP_" + nIndice).ValueParam.Split(',');
            }
            return Usuarios;
        }

        //Inicia el procesamiento de tickets:
        private void ProcesarTicket(Ticket oTicketDatos)
        {
            //Valida campos no vacíos:
            if (!ValidarVacios(oTicketDatos))
            {
                _Funciones.InstanciarFirefoxDriver(ref _oDriver, _cRutaGeckodriver);
                int nIndice = 1;
                for (int i = 0; i < _nIntentosOnBase; i++)
                {
                    UsuariosOnBase(oTicketDatos, nIndice);
                    nIndice += 1;
                    _Funciones.IngresarBPM(_oDriver, _cUrlOnBase, Usuarios[0], Usuarios[1]);
                    if (ExisteElemento("username"))
                    {
                        break;
                    }
                }
                RegistrarBPM(oTicketDatos);
            }
            else
            {
                //Enviar a mesa de control:
                CambiarEstadoTicket(oTicketDatos, _oMesaControl);
            }
        }

        //Valida que no tenga campos vacíos:
        private bool ValidarVacios(Ticket oTicketDatos)
        {
            int[] oCampos = new int[] { eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado };
            return true;
        }

        //Envía el ticket al siguiente estado:
        private void CambiarEstadoTicket(Ticket oTicket, StateAction oAccion)
        {
            //Estado = 1: Mesa de Control, Estado = 2: Notificación de Correo.
            _robot.SaveTicketNextState(oTicket, oAccion.Id);
        }

        private Boolean ExisteElemento(string cIdElemento)
        {
            bool bExiste = true;
            try
            {
                _oDriver.FindElement(By.Id(cIdElemento));
            }
            catch (Exception Ex)
            {
                bExiste = false;
                throw new Exception("Error al conectarse al sistema OnBase", Ex);
            }
            return bExiste;
        }

        private static void SelectByText(SelectElement oElement, string cSearchText)
        {
            var allOptionsThatHaveText = oElement.Options.Where(se => se.Text.Equals(cSearchText, StringComparison.OrdinalIgnoreCase));

            if (allOptionsThatHaveText.Any())
            {
                foreach (var option in allOptionsThatHaveText)
                {
                    option.Click();
                }
                return;
            }

            var optionWithText = oElement.Options.Where(option => option.Text.IndexOf(cSearchText, StringComparison.OrdinalIgnoreCase) >= 0);

            if (optionWithText.Any())
            {
                foreach (var option in optionWithText)
                {
                    option.Click();
                }
                return;
            }

            throw new NoSuchElementException(string.Format("No se encontró el texto: {0} ya sea por coincidencia insensible o porque no está en la colección."));
        }

        //Realiza el registro de anulación en OnBase:
        private void RegistrarBPM(Ticket oTicketDatos)
        {
            try
            {
                LogStartStep(1);
                //Verifica si existe elemento principal:
                //if (!ExisteElemento("username")) return;
                //Clic en pestaña Consultas Personalizadas:
                _oDriver.FindElement(By.Id("controlBarMenuName")).Click();

                //Clic en Nuevo Formulario:
                _Funciones.SeleccionarCombo(_oDriver, "//*[@id='SubMenuOptionsTable']/tbody/tr/td", _cOpcionOnBase);
                //_Funciones.SeleccionarCombo(_oDriver, "SubMenuOptionsTable", "NUEVO FORMULARIO");

                //Seleccionar Solicitud de Pólizas:
                _oDriver.SwitchTo().Frame(_oDriver.FindElement(By.Id("NavPanelIFrame")));
                _oDriver.FindElement(By.XPath("//*[contains(@class,'formName')]")).Click();
                _Funciones.Esperar(2);

                //Frame de Solicitud:
                _oDriver.SwitchTo().DefaultContent();
                _oDriver.SwitchTo().Frame(_oDriver.FindElement(By.Id("frmViewer")));
                _oDriver.SwitchTo().Frame(_oDriver.FindElement(By.Id("uf_hostframe")));

                //Selecciona tipo de solicitud:
                _oDriver.FindElement(By.XPath("//*[@id='tipodesolicitud']/button")).Click();
                _Funciones.Esperar(3);
                _Funciones.SeleccionarCombo(_oDriver, "//body/ul/li/a", _cTipoSolicitud);

                //_oDriver.FindElement(By.XPath("//input[contains(@id,'tipodesolicitud_input')]")).SendKeys(_cTipoSolicitud2);
                //_oDriver.FindElement(By.XPath("//input[contains(@id,'tipodesolicitud_input')]")).SendKeys(Keys.PageDown + Keys.Enter);

                //var oFormulario = new SelectElement(_oDriver.FindElement(By.XPath("//input[contains(@id,'tipodesolicitud_input')]")));
                //SelectByText(oFormulario, "Nuevo formulario");
                //if (_cTipoSolicitud2 == "ENDOSO")
                //{
                //    _oDriver.FindElement(By.XPath("/html/body/ul[1]/li[2]/a")).Click();
                //    string t = _oDriver.FindElement(By.XPath("/html/body/ul[1]/li[2]/a")).Text;
                //}
                //else
                //{
                //    _oDriver.FindElement(By.XPath("/html/body/ul[1]/li[1]/a[contains(@class,'ui-corner-all')]")).Click();
                //}

                //Ingresa fecha hora de email:
                //_oDriver.FindElement(By.Id("fechayhoraderecepción_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_hora_de_email).Value);
                _oDriver.FindElement(By.Id("fechayhoraderecepción_input")).SendKeys("14/08/2018 10:20:15");
                //_oDriver.FindElement(By.Id("fechayhoraderecepción_input")).SendKeys(Keys.Tab);

                _oDriver.FindElement(By.XPath("//*[@id='ui-datepicker-div']/div[3]/button[2]")).Click();

                //Selecciona Línea de Negocio:
                _oDriver.FindElement(By.XPath("//*[@id='linea_negocio']/button")).Click();
                _Funciones.Esperar(2);
                //_Funciones.SeleccionarCombo(_oDriver, "//body/ul[2]/li/a", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_linea).Value);
                _Funciones.SeleccionarCombo(_oDriver, "//body/ul[2]/li/a", "AUTOS");
                _Funciones.Esperar();

                //Selecciona Producto:
                //IList<IWebElement> o= _oDriver.FindElement(By.XPath("//*[@id='producto']/button")).FindElements(By.XPath(("//body/ul[3]/li/a")));
                _oDriver.FindElement(By.XPath("//*[@id='producto']/button")).Click();
                _Funciones.Esperar(2);
                //_Funciones.SeleccionarCombo(_oDriver, "//body/ul[3]/li/a", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.producto).Value);
                _Funciones.SeleccionarCombo(_oDriver, "//body/ul[3]/li/a", "AUTO A MEDIDA");
                _Funciones.Esperar();

                //Seleccionar Tipo de endoso: Necesario validar por el endoso en curso
                _oDriver.FindElement(By.XPath("//*[@id='tipodeendoso']/button")).Click();
                _Funciones.Esperar();
                //_Funciones.SeleccionarCombo(_oDriver, "//body/ul[4]/li/a", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_proceso).Value);
                _Funciones.SeleccionarCombo(_oDriver, "//body/ul[4]/li/a", "ANULACIÓN DE PÓLIZA");
                _Funciones.Esperar(2);

                //_oDriver.FindElement(By.Id("tipodeendoso_input")).SendKeys("ANULACIÓN DE PÓLIZA");
                //_oDriver.FindElement(By.Id("tipodeendoso_input")).SendKeys(Keys.PageDown + Keys.Enter);

                //Seleccionar Motivo de Anulación:
                _oDriver.FindElement(By.XPath("//*[@id='motivo_anulacion']/button")).Click();
                _Funciones.Esperar();
                //_Funciones.SeleccionarCombo(_oDriver, "//body/ul[5]/li/a", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.motivo).Value);
                _Funciones.SeleccionarCombo(_oDriver, "//body/ul[5]/li/a", "DIFICULTADES ECONÓMICAS");
                _Funciones.Esperar();

                //Ingresar número de vehículos y asegurados:
                //_oDriver.FindElement(By.Id("nrounidnroasegnrocertinrobienes_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => 
                //                a.FieldId == (oTicketDatos.TicketValues.FirstOrDefault(b => b.FieldId == eesFields.Default.tipo_de_linea).Value == _cLineaPorDefecto? eesFields.Default.num_vehiculos: eesFields.Default.num_asegurados)).Value);
                _oDriver.FindElement(By.Id("nrounidnroasegnrocertinrobienes_input")).SendKeys("2"); //falta ID de nro asegurados/vehiculos/bienes

                //Ingresa Nro. de Póliza:
                //_oDriver.FindElement(By.Id("vser_nrodepolizasolicitante576_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.numero_de_poliza).Value);
                _oDriver.FindElement(By.Id("vser_nrodepolizasolicitante576_input")).SendKeys("2002925650");

                //Fecha vigencia inicio:
                //_oDriver.FindElement(By.Id("vigenciaendoso_inicio_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.date_inicio_vigencia).Value);
                _oDriver.FindElement(By.Id("vigenciaendoso_inicio_input")).SendKeys("15/08/2018");

                //Fecha vigencia fin:
                //_oDriver.FindElement(By.Id("vigenciaendoso_fin_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.date_fin_vigencia).Value);
                _oDriver.FindElement(By.Id("vigenciaendoso_fin_input")).SendKeys("15/08/2025");

                //Nombre del Contratante:
                //_oDriver.FindElement(By.Id("vser_nombrerazonsocialdelcontratante571_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre_contratante).Value);
                _oDriver.FindElement(By.Id("vser_nombrerazonsocialdelcontratante571_input")).SendKeys("CARLOS REYES");

                //Nombre del Asegurado:
                //_oDriver.FindElement(By.Id("vser_nombrerazonsocialdelcontratante571_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre_asegurado).Value);
                _oDriver.FindElement(By.Id("vser_nombrerazonsocialdelasegurado571_input")).SendKeys("CARLOS REYES");

                //Nro. Doc. Identidad:
                //_oDriver.FindElement(By.Id("nrodedocumentoidentidad_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.numero_de_dni).Value);
                if (ExisteElemento("nrodedocumentoidentidad_input"))
                    //_oDriver.FindElement(By.Id("nrodedocumentoidentidad_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.numero_de_dni).Value);
                    _oDriver.FindElement(By.Id("nrodedocumentoidentidad_input")).SendKeys("41494426");

                //Email de Notificación:
                //_oDriver.FindElement(By.Id("emailnotificacionprincipal_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.email_solicitante).Value);
                _oDriver.FindElement(By.Id("emailnotificacionprincipal_input")).SendKeys("algunos_son_malos@pacificoseguros.com.pe");

                //Adjuntar documentos:
                //_oDriver.FindElement(By.XPath("//*[@id='attach_249']/div[2]/span")).Click();

                _oDriver.FindElement(By.Id("//*[@id='attach_249']/div[2]/span/input")).SendKeys(@"C:\everis.anina");
                _Funciones.Esperar(2);

                //Email de solicitante:
                //_oDriver.FindElement(By.Id("emailnotificacionprincipal_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.email_solicitante).Value);
                if (_oDriver.FindElement(By.Id("emailsolicitante_input")).Text != "") _oDriver.FindElement(By.Id("emailsolicitante_input")).SendKeys("algunos_son_malos@pacificoseguros.com.pe");

                //Guardar y enviar:
                //_oDriver.FindElement(By.XPath("//*[@id='Boton_enviaraemisor']/input")).Click();
                string enviar = _oDriver.FindElement(By.XPath("//*[@id='Boton_enviaraemisor']/input")).Text;
                LogEndStep(1);
            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
                throw new Exception(Ex.Message, Ex);
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
