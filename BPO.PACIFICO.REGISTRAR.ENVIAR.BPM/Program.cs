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
        private static BaseRobot<Program> _oRobot = null;
        private static IWebDriver _oDriver = null;
        private static int _nIdEstadoError;
        private static int _nIdEstadoFinal;
        private static string _cUrlOnBase = string.Empty;
        private static string _cTipoSolicitud = string.Empty;
        private static int _nIntentosOnBase;
        private static string _cOpcionFormulario = string.Empty;
        private static string _cOpcionWorkflow = string.Empty;
        private static string _cLineaPorDefecto = string.Empty;
        private static string _cLineaTicket = string.Empty;
        private static string _cRutaGeckodriver = string.Empty;
        private static string _cRutaFirefox = string.Empty;
        private static string _cRutaAdjuntos = string.Empty;
        private static string _cElemento = string.Empty;
        private static string _cEndosoAnulacion = string.Empty;
        private static string _cBPMWebDriver = string.Empty;
        private static string _cGeckodriver = string.Empty;
        private static StateAction _oMesaControl;
        private static StateAction _oRegistro;
        private static Functions _Funciones;
        private static string[] Usuarios;
        #endregion

        static void Main(string[] args)
        {
            try
            {
                _oRobot = new BaseRobot<Program>(args);
                _Funciones = new Functions();
                _oRobot.Start();
            }
            catch (Exception Ex) { Console.WriteLine(Ex.Message); }
        }

        protected override void Start()
        {
            if (_oRobot.Tickets.Count < 1)
                return;

            ObtenerParametros();
            LogStartStep(4);
            foreach (Ticket oTicket in _oRobot.Tickets)
            {
                try
                {
                    _oMesaControl = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdEstadoError);
                    _oRegistro = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdEstadoFinal);
                    _cLineaTicket = _Funciones.ObtenerValorDominio(oTicket, Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.linea).Value));
                    ProcesarTicket(oTicket);
                }
                catch (Exception Ex)
                {
                    CambiarEstadoTicket(oTicket, _oMesaControl, Ex.Message);
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
                _nIdEstadoError = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoError").ValueParam);
                _nIdEstadoFinal = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoFinal").ValueParam);
                _cUrlOnBase = _oRobot.GetValueParamRobot("URLOnBase").ValueParam;
                _cTipoSolicitud = _oRobot.GetValueParamRobot("TipoSolicitud").ValueParam;
                _nIntentosOnBase = Convert.ToInt32(_oRobot.GetValueParamRobot("AccesoOnBase_Intentos").ValueParam);
                _cOpcionFormulario = _oRobot.GetValueParamRobot("OnBase_Formulario").ValueParam;
                _cOpcionWorkflow = _oRobot.GetValueParamRobot("OnBase_Workflow").ValueParam;
                _cLineaPorDefecto = _oRobot.GetValueParamRobot("LineaPorDefecto").ValueParam;
                _cRutaGeckodriver = _oRobot.GetValueParamRobot("Ruta_geckodriver").ValueParam;
                _cRutaFirefox = _oRobot.GetValueParamRobot("Ruta_Firefox").ValueParam;
                _cRutaAdjuntos = _oRobot.GetValueParamRobot("Ruta_AdjuntosBPM").ValueParam;
                _cEndosoAnulacion = _oRobot.GetValueParamRobot("EndosoAnulacion").ValueParam;
                _cBPMWebDriver = _oRobot.GetValueParamRobot("BPMWebDriver").ValueParam;
                _cGeckodriver = _oRobot.GetValueParamRobot("Geckodriver").ValueParam;
            }
            catch (Exception Ex) { LogFailStep(12, Ex); }
        }

        //Obtiene los usuarios OnBase según Línea de negocio:
        private string[] UsuariosOnBase(int nIndice = 1)
        {
            try
            {
                if (_cLineaTicket == _cLineaPorDefecto)
                {
                    //Usuario y contraseña de Autos:
                    Usuarios = _oRobot.GetValueParamRobot("AccesoOnBase_Autos_" + nIndice).ValueParam.Split(',');
                }
                else
                {
                    //Usuario y contraseña de Riesgos Generales y Líneas Personales:
                    Usuarios = _oRobot.GetValueParamRobot("AccesoOnBase_LLPP_" + nIndice).ValueParam.Split(',');
                }
                return Usuarios;
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al obtener datos de usuario: " + Ex.Message, Ex); }
        }

        //Inicia el procesamiento de tickets:
        private void ProcesarTicket(Ticket oTicketDatos)
        {
            //Valida campos no vacíos:
            if (ValidarVacios(oTicketDatos))
            {
                int nIndice = 1;
                for (int i = 0; i < _nIntentosOnBase; i++)
                {
                    _Funciones.InstanciarFirefoxDriver(ref _oDriver, _cRutaGeckodriver, _cRutaFirefox, _cBPMWebDriver, _cGeckodriver);
                    UsuariosOnBase(nIndice);
                    _Funciones.IngresarBPM(_oDriver, _cUrlOnBase, Usuarios[0], Usuarios[1]);
                    if (_Funciones.ExisteElemento(_oDriver, "controlBarMenuName", _nIntentosOnBase))
                    {
                        break;
                    }
                    _Funciones.CerrarDriver(_oDriver);
                    nIndice += 1;
                }
                RegistrarBPM(oTicketDatos);
            }
            else
            {
                //Enviar a mesa de control:
                CambiarEstadoTicket(oTicketDatos, _oMesaControl, "El ticket " + Convert.ToString(oTicketDatos.Id) + " no cuenta con todos los datos necesarios.");
            }
        }

        //Valida que no tenga campos vacíos:
        private Boolean ValidarVacios(Ticket oTicketDatos)
        {
            try
            {
                int[] oCampos = new int[] { eesFields.Default.cuenta_nombre, eesFields.Default.asegurado_nombre,
                                            eesFields.Default.email_solicitante, eesFields.Default.fecha_hora_de_email,
                                            eesFields.Default.poliza_fec_ini_vig, eesFields.Default.poliza_fec_fin_vig};

                return _Funciones.ValidarCamposVacios(oTicketDatos, oCampos);
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al validar campos del Ticket: " + Convert.ToString(oTicketDatos.Id), Ex); }
        }

        //Envía el ticket al siguiente estado:
        private void CambiarEstadoTicket(Ticket oTicket, StateAction oAccion, string cMensaje = "")
        {
            _oRobot.SaveTicketNextState(cMensaje == "" ? oTicket : _Funciones.MesaDeControl(oTicket, cMensaje), oAccion.Id);
        }

        //Realiza el registro de anulación en OnBase:
        private void RegistrarBPM(Ticket oTicketDatos)
        {
            try
            {
                LogStartStep(1);
                IniciaFormularioBPM();
                LlenarFormularioBPM(oTicketDatos);
            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
                throw new Exception(Ex.Message + " :" + _cElemento, Ex);
            }
        }

        private void IniciaFormularioBPM()
        {
            //Clic en pestaña Consultas Personalizadas:
            if (!_Funciones.ExisteElemento(_oDriver, "controlBarMenuName", _nIntentosOnBase))
                IniciaFormularioBPM();

            _cElemento = _oDriver.FindElement(By.XPath("//*[@id ='DropDownContainer']/tbody/tr/td[2]")).Text;
            _oDriver.FindElement(By.XPath("//*[@id ='DropDownContainer']/tbody/tr/td[2]")).Click();

            if (_oDriver.FindElement(By.XPath("/html/body/div[8]")).Displayed)
            {
                //Clic en Nuevo Formulario:
                _Funciones.SeleccionarListBox(_oDriver, "//*[@id='SubMenuOptionsTable']/tbody/tr", _cOpcionFormulario);

                //Seleccionar Solicitud de Pólizas:
                _oDriver.SwitchTo().Frame(_oDriver.FindElement(By.Id("NavPanelIFrame")));
                _oDriver.FindElement(By.XPath("//html/body/form/div[2]/div[2]")).Click();
                _cElemento = _oDriver.FindElement(By.XPath("//html/body/form/div[2]/div[2]")).Text;
                _Funciones.Esperar(8);
            }
            else
            {
                IniciaFormularioBPM();
            }
        }

        private void LlenarFormularioBPM(Ticket oTicketDatos)
        {
            //Frame de Solicitud:
            _cElemento = "Panel de solicitud";
            _oDriver.SwitchTo().DefaultContent();
            _oDriver.SwitchTo().Frame(_oDriver.FindElement(By.Id("frmViewer")));
            _oDriver.SwitchTo().Frame(_oDriver.FindElement(By.Id("uf_hostframe")));

            //Ingresa fecha hora de email:
            _cElemento = "Fecha y hora de solicitud";
            _oDriver.FindElement(By.Id("fechayhoraderecepción_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_hora_de_email).Value);
            //_oDriver.FindElement(By.Id("fechayhoraderecepción_input")).SendKeys("14/08/2018 10:20:15" + Keys.Tab);

            //Selecciona tipo de solicitud:
            _cElemento = "Tipo de solicitud";
            _oDriver.FindElement(By.XPath("//*[@id='tipodesolicitud']/button")).Click();
            _Funciones.Esperar();
            _Funciones.SeleccionarListBox(_oDriver, "/html/body/ul[1]/li", _cTipoSolicitud);

            //Selecciona Línea de Negocio:
            _oDriver.FindElement(By.XPath("//*[@id='linea_negocio']/button")).Click();
            _Funciones.Esperar();
            _cElemento = "Línea de Negocio";
            string cLinea = _Funciones.ObtenerValorDominio(oTicketDatos, Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.linea).Value));
            _Funciones.SeleccionarListBox(_oDriver, "//body/ul[2]/li/a", cLinea);

            //Selecciona Producto:
            _oDriver.FindElement(By.XPath("//*[@id='producto']/button")).Click();
            _Funciones.Esperar(2);
            _cElemento = "Producto";
            _Funciones.SeleccionarListBox(_oDriver, "//body/ul[3]/li/a", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.producto).Value);
            //_Funciones.SeleccionarListBox(_oDriver, "//body/ul[3]/li/a", "AUTO A MEDIDA");            
            _Funciones.Esperar(3);

            //Seleccionar Tipo de endoso:
            _cElemento = "Tipo de Endoso";
            _oDriver.FindElement(By.XPath("//*[@id='tipodeendoso']/button")).Click();
            _Funciones.Esperar();
            string cProceso = _Funciones.ObtenerValorDominio(oTicketDatos, Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_proceso).Value));
            _Funciones.SeleccionarListBox(_oDriver, "//body/ul[4]/li/a", cProceso);
            _Funciones.Esperar(3);

            //Seleccionar Motivo de Anulación:
            if (_oDriver.FindElement(By.Id("motivo_anulacion")).Displayed)
            {
                _cElemento = "Motivo de Anulación";
                _oDriver.FindElement(By.XPath("//*[@id='motivo_anulacion']/button")).Click();
                _Funciones.Esperar();
                string cMotivo = _Funciones.ObtenerValorDominio(oTicketDatos, Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.motivo_anular).Value));
                _Funciones.SeleccionarListBox(_oDriver, "//body/ul[5]/li/a", cMotivo);
            }

            //Ingresar número de vehículos y asegurados:
            _cElemento = "Número Unidades/Asegurados/Bienes";
            _oDriver.FindElement(By.Id("nrounidnroasegnrocertinrobienes_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a =>
                            a.FieldId == (oTicketDatos.TicketValues.FirstOrDefault(b => b.FieldId == eesFields.Default.linea).Value == _cLineaPorDefecto ? 1 : 1)).Value);
            //_oDriver.FindElement(By.Id("nrounidnroasegnrocertinrobienes_input")).SendKeys("2");           

            //Ingresa Nro. de Póliza:
            _cElemento = "Número de Póliza";
            _oDriver.FindElement(By.Id("vser_nrodepolizasolicitante576_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.numero_de_poliza).Value);
            //_oDriver.FindElement(By.Id("vser_nrodepolizasolicitante576_input")).SendKeys("2002925650");

            //Fecha vigencia inicio:
            _cElemento = "Fecha Inicio";
            _oDriver.FindElement(By.Id("vigenciaendoso_inicio_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_fec_ini_vig).Value);
            //_oDriver.FindElement(By.Id("vigenciaendoso_inicio_input")).SendKeys("15/08/2018");

            //Fecha vigencia fin:
            _cElemento = "Fecha Fin";
            _oDriver.FindElement(By.Id("vigenciaendoso_fin_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_fec_fin_vig).Value + Keys.Tab);
            //_oDriver.FindElement(By.Id("vigenciaendoso_fin_input")).SendKeys("15/08/2025" + Keys.Tab);            

            //Nombre del Contratante:
            _oDriver.FindElement(By.Id("vser_nombrerazonsocialdelcontratante571_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.cuenta_nombre).Value.ToUpper());
            //_oDriver.FindElement(By.Id("vser_nombrerazonsocialdelcontratante571_input")).SendKeys("CARLOS REYES");
            _cElemento = "Nombre Contratante";

            //Nombre del Asegurado:
            _cElemento = "Nombre Asegurado";
            _oDriver.FindElement(By.Id("vser_nombrerazonsocialdelcontratante571_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.asegurado_nombre).Value.ToUpper());
            //_oDriver.FindElement(By.Id("vser_nombrerazonsocialdelasegurado571_input")).SendKeys("CARLOS REYES");

            //Nro. Doc. Identidad:
            if (_oDriver.FindElement(By.Id("nrodedocumentoidentidad")).Displayed)
            {
                _cElemento = "Documento de identidad";
                _oDriver.FindElement(By.Id("nrodedocumentoidentidad_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.numero_de_dni).Value);
                //_oDriver.FindElement(By.Id("nrodedocumentoidentidad_input")).SendKeys("41494426");
            }

            //Email de Notificación:
            _cElemento = "Email Principal para notificación";
            _oDriver.FindElement(By.Id("emailnotificacionprincipal_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.email_solicitante).Value);
            //_oDriver.FindElement(By.Id("emailnotificacionprincipal_input")).SendKeys("algunos_son_malos@pacificoseguros.com.pe");

            //Adjuntar documentos:
            _cElemento = "Adjuntar documentos";
            foreach (string cDocumentos in Adjuntos(oTicketDatos))
            {
                _oDriver.FindElement(By.XPath("//*[@id='attach_251']/div[2]/span/input")).SendKeys(cDocumentos);
            }
            _Funciones.Esperar(2);

            //Email de solicitante:
            _cElemento = "Email de solicitante";
            if (_oDriver.FindElement(By.Id("emailsolicitante_input")).Text.Length == 0) _oDriver.FindElement(By.Id("emailsolicitante_input")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.email_solicitante).Value);
            //if (_oDriver.FindElement(By.Id("emailsolicitante_input")).Text.Length == 0) _oDriver.FindElement(By.Id("emailsolicitante_input")).SendKeys("algunos_son_malos@pacificoseguros.com.pe");

            //Guardar y enviar:
            //_oDriver.FindElement(By.XPath("//*[@id='Boton_enviaraemisor']/input")).SendKeys(Keys.Enter);
            _Funciones.Esperar(3);

            if (_Funciones.VerificarRegistroBPM(_oDriver))
            {
                _Funciones.Esperar(7);
                WorkflowOnBase();
                CambiarEstadoTicket(oTicketDatos, _oRegistro);
                LogEndStep(1);
            }
            else
            {
                //Repetir operación:
                _oDriver.SwitchTo().DefaultContent();
                _oDriver.SwitchTo().Frame(_oDriver.FindElement(By.Id("NavPanelIFrame")));
                _oDriver.FindElement(By.XPath("//html/body/form/div[2]/div[2]")).Click();
                _oDriver.SwitchTo().Alert().Accept();
                _Funciones.Esperar(10);
                LlenarFormularioBPM(oTicketDatos);
            }
        }

        //Obtiene documentos adjuntos según endoso:
        private string[] Adjuntos(Ticket oTicketDatos)
        {
            try
            {
                return oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.documentos).Value.Split(',');
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al obtener documentos adjuntos: " + Ex.Message, Ex); }
        }

        private void WorkflowOnBase()
        {
            try
            {
                _cElemento = _oDriver.FindElement(By.XPath("//*[@id ='MainMenuHeader']/tbody/tr/td[3]")).Text;
                _oDriver.FindElement(By.XPath("//*[@id ='MainMenuHeader']/tbody/tr/td[3]")).Click();

                if (_oDriver.FindElement(By.XPath("/html/body/div[7]")).Displayed)
                {
                    //Clic en Workflow:
                    _Funciones.SeleccionarListBox(_oDriver, "//*[@id='SubMenuOptionsTable']/tbody/tr", _cOpcionWorkflow);
                    _Funciones.Esperar(5);

                    //Carga ventana de Workflow en nueva instancia Firefox.

                    //En pestaña Carpeta de Trabajo:

                    //Se muestra grilla con lista de documentos (panel derecho), es necesario actualizar grilla.

                    //Seleccionar el documento según Nro. de Trámite (El que se genera al registrar formulario BPM).

                    _Funciones.Esperar(5);

                    //Aprobar o rechazar trámite:
                    _Funciones.Esperar(2);

                    //Seleccionar canal: Por descripción, validar de qué forma se obtendrá el nombre de canal desde PolicyCenter.

                    //Seleccionar Agente: Ingresando el número de agente sin ceros.

                    //Seleccionar motivo de rechazo emisor: Validar desde dónde se obtiene el motivo.

                    //Ingresar Comentario Rechazo Emisor: Validar si este dato es obligatorio y opcional y de dónde se obtiene.

                    //Finalmente, presionar botón Grabar.

                    _Funciones.Esperar(4);
                    //Seleccionar opción NO (Enviar notificación).         
                }
                else
                {
                    WorkflowOnBase();
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al ingresar a la opción Workflow: " + Ex.Message, Ex); }
        }
    }
}
