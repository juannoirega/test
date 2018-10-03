using everis.Ees.Proxy;
using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using Robot.Util.Nacar;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BPO.PACIFICO.ACTUALIZAR.DATOS.CLIENTE
{
    class Program : IRobot
    {
        #region VARIABLES Y PARÁMETROS
        private static BaseRobot<Program> _oRobot = null;
        private static IWebDriver _driverGlobal = null;
        private static string campo = string.Empty;
        private static string texto = string.Empty;
        static string[] _valoresTickets = new string[10];
        static string[] _valoresTickets_Ident = new string[10];
        private string _cUrlContactManager = string.Empty;
        private string _usuarioContactManager = string.Empty;
        private string _contraseñaContactManager = string.Empty;
        private static int _nIdEstadoError; //Mesa de Control
        private static int _nIdEstadoSiguiente; //Robot Crear Ticket Hijo
        private static int _nIntentosPolicyCenter;
        private static string _cUrlPolicyCenter = string.Empty;
        private static int _nFieldId;
        private static string _cFieldLabel = string.Empty;
        private static int _nFormulario;
        private static string[] _TipoContacto;
        private static string[] _Usuarios;
        private static string[] _DominioLineas;
        private static string[] _DominioComplejidad;
        private static string[] _DominioMotivo;
        private static string[] _TiempoEspera;
        private static int _nIndice;
        private static string _cElemento = string.Empty;
        private static string _cLineaAutos = string.Empty;
        private static string _cLineaRRGG = string.Empty;
        private static string _cLineaAlianzas = string.Empty;
        private static string _cLineaLLPP = string.Empty;
        private static string _cLinea = string.Empty;
        private static string _cNombreOferta = string.Empty;
        private static string _cCeldaLimitante = string.Empty;
        private static string _cOrdenTrabajo = string.Empty;
        private static int _nIndexFilaCuenta = 0;
        private static int _nIndexCeldaCuenta = 0;
        private static bool _bControl = false;
        private static int _reprocesoContador = 0;
        private static int _idEstadoRetorno = 0;
        private static string _cTicketValue = string.Empty;
        private static Functions _Funciones;
        private static StateAction _oMesaControl;
        private static StateAction _oTicketHijo;
        #endregion

        static void Main(string[] args)
        {
            try
            {
                _Funciones = new Functions();
                _oRobot = new BaseRobot<Program>(args);
                _oRobot.Start();
            }
            catch (Exception Ex) { Console.WriteLine(Ex.Message); }
        }

        protected override void Start()
        {
            if (_oRobot.Tickets.Count < 1)
                return;

            Inicio();
            ObtenerParametros();
            LogStartStep(4);
            foreach (Ticket oTicket in _oRobot.Tickets)
            {
                try
                {
                    var valoresReprocesamiento = _Funciones.ObtenerValoresReprocesamiento(oTicket);
                    if (valoresReprocesamiento.Count > 0) { _reprocesoContador = valoresReprocesamiento[0]; _idEstadoRetorno = valoresReprocesamiento[1]; }
                    _oMesaControl = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdEstadoError);
                    _oTicketHijo = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdEstadoSiguiente);
                    //Obteniendo Línea de Negocio:
                    _cLinea = _Funciones.GetDomainValue(Convert.ToInt32(_DominioLineas[0]), Convert.ToInt32(_DominioLineas[1]), Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.idlinea).Value)).ToUpperInvariant();
                    ProcesarTicket(oTicket);
                }
                catch (Exception Ex)
                {
                    _reprocesoContador++;
                    _Funciones.GuardarIdPlantillaNotificacion(oTicket,
                        Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.idproceso).Value),
                        Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.idlinea).Value),
                        false
                        );
                    _Funciones.GuardarValoresReprocesamiento(oTicket, _reprocesoContador, _idEstadoRetorno);
                    CambiarEstadoTicket(oTicket, _oMesaControl, Ex.Message);
                    LogFailStep(30, Ex);
                }
                finally { _Funciones.CerrarDriver(_driverGlobal); }
            }
        }

        private void Inicio()
        {
            Console.WriteLine("♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦ ROBOT ♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦");
            Console.WriteLine("           Robot Actualizar Datos del Cliente          ");
            Console.WriteLine("♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦");
        }

        //Obtiene los parámetros asociados al robot:
        private void ObtenerParametros()
        {
            try
            {
                _cUrlContactManager = _oRobot.GetValueParamRobot("URLContactManager").ValueParam;
                _usuarioContactManager = _oRobot.GetValueParamRobot("UsuarioContactManager").ValueParam;
                _contraseñaContactManager = _oRobot.GetValueParamRobot("PasswordContactManager").ValueParam;
                _nIdEstadoError = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoError").ValueParam);
                _nIdEstadoSiguiente = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoSiguiente").ValueParam);
                _nIntentosPolicyCenter = Convert.ToInt32(_oRobot.GetValueParamRobot("nIntentosPolicyCenter").ValueParam);
                _cUrlPolicyCenter = _oRobot.GetValueParamRobot("URLPolicyCenter").ValueParam;
                _TipoContacto = _oRobot.GetValueParamRobot("TipoContacto").ValueParam.Split(',');
                _cLineaAutos = _oRobot.GetValueParamRobot("LineaAutos").ValueParam;
                _cLineaRRGG = _oRobot.GetValueParamRobot("LineaRRGG").ValueParam;
                _cLineaAlianzas = _oRobot.GetValueParamRobot("LineaAlianzas").ValueParam;
                _cLineaLLPP = _oRobot.GetValueParamRobot("LineaLLPP").ValueParam;
                _cCeldaLimitante = _oRobot.GetValueParamRobot("CeldaLimitante").ValueParam;
                _nIndexFilaCuenta = Convert.ToInt32(_oRobot.GetValueParamRobot("IndexFilaCuenta").ValueParam);
                _nIndexCeldaCuenta = Convert.ToInt32(_oRobot.GetValueParamRobot("IndexCeldaCuenta").ValueParam);
                _DominioLineas = _oRobot.GetValueParamRobot("ParametrosDominioLineas").ValueParam.Split(',');
                _DominioComplejidad = _oRobot.GetValueParamRobot("ParametrosDominioComplejidad").ValueParam.Split(',');
                _DominioMotivo = _oRobot.GetValueParamRobot("ParametrosDominioMotivo").ValueParam.Split(',');
                _TiempoEspera = _oRobot.GetValueParamRobot("TiempoEspera").ValueParam.Split(',');
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al obtener parámetros del Robot: " + Ex.Message, Ex); }
        }

        //Envía el ticket al siguiente estado:
        private void CambiarEstadoTicket(Ticket oTicket, StateAction oAccion, string cMensaje = "")
        {
            _oRobot.SaveTicketNextState(cMensaje == "" ? oTicket : _Funciones.MesaDeControl(oTicket, cMensaje), oAccion.Id);
        }

        private void ProcesarTicket(Ticket oTicket)
        {
            ContactManager(oTicket);
            PolicyCenter(oTicket);
        }

        private void IniciarSistema(string cUrl)
        {
            _nIndice = 1;
            for (int i = 0; i < _nIntentosPolicyCenter; i++)
            {
                Credenciales();
                _Funciones.AbrirSelenium(ref _driverGlobal);
                if (_Funciones.StartSystem(_driverGlobal, cUrl, By.Id("diagnose"), _Usuarios, _nIntentosPolicyCenter)) { break; }
                _nIndice += 1;
            }
        }

        #region CONTACT MANAGER
        private void ContactManager(Ticket oTicket)
        {
            IniciarSistema(_cUrlContactManager);
            ActualizarContactManager(oTicket);
        }

        private void ActualizarContactManager(Ticket oTicketDatos)
        {
            try
            {
                //Opteniendo DNI
                _valoresTickets[0] = oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nro_dni).Value;
                //Opteniendo RUC
                _valoresTickets[1] = oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nro_ruc).Value;

                if (_valoresTickets[0].Length > 0) { Filtros(_valoresTickets[0]); }
                else if (_valoresTickets[1].Length > 0) { Filtros(_valoresTickets[1], false); }

                if (!_Funciones.ExisteElemento(_driverGlobal, By.XPath("//*[@id='ABContactSearch:ABContactSearchScreen:_msgs_msgs']/div")))
                {
                    AccederRegistro();
                    GetFieldIdByNames(oTicketDatos, eesFields.Default.listacampos);
                    FinalizarActualizacion(oTicketDatos);
                }
                else { CambiarEstadoTicket(oTicketDatos, _oMesaControl, "No se encontraron registros para " + _valoresTickets[0] == "" ? _valoresTickets[1] : _valoresTickets[0]); }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al actualizar en Contact Manager: " + Ex.Message, Ex); }
        }

        public void Filtros(string valor, bool bPersona = true)
        {
            try
            {
                if (bPersona)
                {
                    IngresarTextoSelect("ABContactSearch:ABContactSearchScreen:ContactSearchDV:ContactSubtype", _TipoContacto[0]);
                    IngresarTextoSelect("ABContactSearch:ABContactSearchScreen:ContactSearchDV:PrimaryOfficialIDTypeExt", _TipoContacto[1]);
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                    _driverGlobal.FindElement(By.Id("ABContactSearch:ABContactSearchScreen:ContactSearchDV:TaxID")).SendKeys(Keys.Control + "e");
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                    _driverGlobal.FindElement(By.Id("ABContactSearch:ABContactSearchScreen:ContactSearchDV:TaxID")).SendKeys(valor);
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                    _driverGlobal.FindElement(By.ClassName("bigButton_link")).Click();
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                }
                else
                {
                    IngresarTextoSelect("ABContactSearch:ABContactSearchScreen:ContactSearchDV:ContactSubtype", _TipoContacto[2]);
                    IngresarTextoSelect("ABContactSearch:ABContactSearchScreen:ContactSearchDV:PrimaryOfficialIDTypeExt", _TipoContacto[3]);
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                    _driverGlobal.FindElement(By.Id("ABContactSearch:ABContactSearchScreen:ContactSearchDV:TaxID")).SendKeys(Keys.Control + "e");
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                    _driverGlobal.FindElement(By.Id("ABContactSearch:ABContactSearchScreen:ContactSearchDV:TaxID")).SendKeys(valor);
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                    _driverGlobal.FindElement(By.ClassName("bigButton_link")).Click();
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error: " + Ex.Message, Ex); }
        }

        public void AccederRegistro()
        {
            _nFormulario = 1;
            try
            {
                _driverGlobal.FindElement(By.Id("ABContactSearch:ABContactSearchScreen:ContactSearchResultsLV:0:DisplayName")).Click();
                _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                _driverGlobal.FindElement(By.XPath("//*[@id='ContactDetail:ABContactDetailScreen:ContactBasicsDV_tb:Edit']/span[2]")).Click();
                _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error: " + Ex.Message, Ex); }
        }

        private void GetFieldIdByNames(Ticket oTicketDatos, int nIdListField)
        {
            try
            {
                var container = ODataContextWrapper.GetContainer();
                String[] oFieldNames = oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == nIdListField).Value.Split(',');

                foreach (string cField in oFieldNames)
                {
                    _cFieldLabel = container.Fields.Where(f => f.Name == cField).Select(f => new { f.Label }).FirstOrDefault().Label;
                    _nFieldId = container.Fields.Where(f => f.Name == cField).Select(f => new { f.Id }).FirstOrDefault().Id;
                    _cTicketValue = oTicketDatos.TicketValues.FirstOrDefault(t => t.FieldId == Convert.ToInt32(_nFieldId)).Value;

                    if (_Funciones.IsFieldEdit(oTicketDatos, _nFieldId))
                    {
                        if (_nFormulario == 1) { EditarFormulario(_cFieldLabel, _cTicketValue); }
                        else if (_nFormulario == 2) { FormularioEditarPoliza(_nFieldId); }
                        else { FormularioEditarCuenta(_nFieldId); }
                    }
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error: " + Ex.Message, Ex); }
        }

        public void EditarFormulario(String campo, String texto)
        {
            try
            {
                switch (campo)
                {
                    case "Nacionalidad":
                        CambiarNacionalidad(texto);
                        break;
                    case "País de procedencia":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:CountryOfOriginExt", texto);
                        break;
                    case "Nombre (s)":
                    case "Razón Social":
                        EscribirElementoXPathActualizarDatos(0, 4, texto, "textBox");
                        break;
                    case "Apellido Paterno":
                    case "Nombre comercial":
                        EscribirElementoXPathActualizarDatos(0, 5, texto, "textBox");
                        break;
                    case "Apellido Materno":
                        EscribirElementoXPathActualizarDatos(0, 6, texto, "textBox");
                        break;
                    case "Nombre corto":
                        EscribirElementoXPathActualizarDatos(0, 8, texto, "textBox");
                        break;
                    case "Prefijo":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:Prefix", texto);
                        break;
                    case "Teléfono principal":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:ABPhoneDetailsInputSet:PrimaryPhone", texto);
                        break;
                    case "País del teléfono":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:ABPhoneDetailsInputSet:HomePhoneCountry", texto);
                        break;
                    case "Indicativo(código de área)":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:ABPhoneDetailsInputSet:HomeAreaCodeExtPeru", texto);
                        break;
                    case "Teléfono de Casa":
                        EscribirElementoXPathActualizarDatos(0, 18, texto, "textBox_error");
                        break;
                    case "Dirección principal de correo electrónico":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:notVendor:PrimaryEmailTypeExt", texto);
                        break;
                    case "Correo Personal":
                        EscribirElementoXPathActualizarDatos(0, 41, texto, "textBox");
                        break;
                    case "País":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_Country", texto);
                        break;
                    case "Departamento":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_Country", texto);
                        break;
                    case "Provincia":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_Province", texto);
                        break;
                    case "Distrito":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:abc", texto);
                        break;
                    case "Tipo de calle":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_StreetType", texto);
                        break;
                    case "Nombre de la calle":
                        EscribirElementoXPathActualizarDatos(0, 53, texto, "textBox");
                        break;
                    case "Número":
                        EscribirElementoXPathActualizarDatos(0, 54, texto, "textBox");
                        break;
                    case "Referencia":
                        EscribirElementoXPathActualizarDatos(0, 55, texto, "textBox");
                        break;
                    case "Fecha de nacimiento":
                        texto = texto.Replace("/", "");
                        EscribirElementoXPathActualizarDatos(5, 5, texto, "textBox");
                        break;
                    case "Sexo":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:ABPersonVendorInputSet:Gender", texto);
                        break;
                    case "Estado civil":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:ABPersonVendorInputSet:MaritalStatus", texto);
                        break;
                    case "Fecha de inicio de actividades":
                        texto = texto.Replace("/", "");
                        EscribirElementoXPathActualizarDatos(6, 5, texto, "textBox");
                        break;
                    case "Actividad económica":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:EconomicSectorActivityInputSet:EconomicSubSectorExt", texto);
                        break;
                    case "Sector Económico":
                        IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:EconomicSectorActivityInputSet:EconomicSectorExt", texto);
                        break;
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al editar formulario: " + Ex.Message, Ex); }
        }

        public void EscribirElementoXPathActualizarDatos(int index, int posicion, string texto, string clase)
        {
            //IWebElement _element = null; ;
            string xPath = "//*[@id='ContactDetail:ABContactDetailScreen:ContactBasicsDV:" + index + "']/tbody/tr[" + posicion + "]/td[5]/input[@class='" + clase + "']";
            _driverGlobal.FindElement(By.XPath(xPath)).SendKeys("");
            _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
            _driverGlobal.FindElement(By.XPath(xPath)).SendKeys(Keys.Control + "e");
            _driverGlobal.FindElement(By.XPath(xPath)).SendKeys(texto);
        }

        public void CambiarNacionalidad(string Texto)
        {
            if (Texto == "Peruano (a)")
                _driverGlobal.FindElement(By.XPath("//*[@id='ContactDetail: ABContactDetailScreen:ContactBasicsDV: NationalityExt_N']")).Click();
            else
                _driverGlobal.FindElement(By.XPath("//*[@id='ContactDetail:ABContactDetailScreen:ContactBasicsDV:NationalityExt_E']")).Click();
        }

        public void IngresarTextoSelect(String name, String texto)
        {
            SelectElement elemen = new SelectElement(_driverGlobal.FindElement(By.Name(name)));
            elemen.SelectByText(texto);
            _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
        }

        private void FinalizarActualizacion(Ticket oTicketDatos)
        {
            //Guardar Cambios
            _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
            _driverGlobal.FindElement(By.XPath("//*[@id='ContactDetail:ABContactDetailScreen:ContactBasicsDV_tb:Update']/span[2]")).Click();
            if (_Funciones.ExisteElemento(_driverGlobal, By.Id("WebMessageWorksheet:WebMessageWorksheetScreen:grpMsgs_msgs")))
            {
                CambiarEstadoTicket(oTicketDatos, _oMesaControl, "Ocurrió un error: " + _driverGlobal.FindElement(By.Id("WebMessageWorksheet:WebMessageWorksheetScreen:grpMsgs_msgs")).Text);
                _driverGlobal.FindElement(By.XPath("//*[@id='ContactDetail:ABContactDetailScreen:ContactBasicsDV_tb:Cancel']/span[2]")).Click();
            }
            _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
            _Funciones.CerrarDriver(_driverGlobal);
        }
        #endregion

        #region POLICYCENTER
        private void PolicyCenter(Ticket oTicket)
        {
            IniciarSistema(_cUrlPolicyCenter);
            ActualizarPolicyCenter(oTicket);

            if (String.IsNullOrWhiteSpace(_cOrdenTrabajo))
            {
                CambiarEstadoTicket(oTicket, _oMesaControl, _bControl == true ? "La fecha efectiva del cambio no se encuentra dentro del rango de vigencia." : "Ocurrió un error en Análisis de Riesgos para la línea " + _cLinea + ", se requiere aprobación del endoso.");
            }
            else
            {
                AgregarValoresTicket(oTicket);
                _Funciones.GuardarIdPlantillaNotificacion(oTicket,
                    Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.idproceso).Value),
                    Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.idlinea).Value)
                    );
                //Si todo es conforme, pasa al estado Crear Ticket Hijo:
                if (_reprocesoContador > 0) { _reprocesoContador = 0; _idEstadoRetorno = 0; _Funciones.GuardarValoresReprocesamiento(oTicket, _reprocesoContador, _idEstadoRetorno); }
                CambiarEstadoTicket(oTicket, _oTicketHijo);
            }
            LogEndStep(4);
        }

        //Obtiene los usuarios con sus respectivas contraseñas:
        private string[] Credenciales()
        {
            try
            {
                //Usuario y contraseña:
                _Usuarios = _oRobot.GetValueParamRobot("AccesoPCyCM_" + _nIndice).ValueParam.Split(',');
                return _Usuarios;
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al obtener datos de usuario: " + Ex.Message, Ex); }
        }

        private void ActualizarPolicyCenter(Ticket oTicketDatos)
        {
            try
            {
                _nFormulario = 2;
                LogStartStep(1);

                if (_cLinea == _cLineaAutos)
                {
                    //Busca datos de Póliza por Nro. de Póliza:
                    _cElemento = "Buscar póliza";
                    _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_nro).Value);

                    //Obtener nombre de la oferta:
                    _cNombreOferta = _Funciones.FindElement(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_PolicyDV:Offering"), Convert.ToInt32(_TiempoEspera[3])).Text;

                    IniciarCambioPoliza(oTicketDatos);
                    if (!FormularioCambioPoliza(oTicketDatos)) { _bControl = true; return; }
                    if (!String.IsNullOrWhiteSpace(_cNombreOferta)) { SeleccionarOferta(oTicketDatos); }

                    //Método para actualizar datos:
                    GetFieldIdByNames(oTicketDatos, eesFields.Default.listacampos);

                    if (AnalisisDeRiesgos())
                    {
                        ConfirmarTrabajo();
                    }
                    else
                    {
                        //Cancelar cotización:
                        _cElemento = "Cancelar cotización";
                        _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:JobWizardToolbarButtonSet:WithdrawJob")).Click();
                        _Funciones.VerificarVentanaAlerta(_driverGlobal);
                        _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[2]));
                    }
                }
                else if (_cLinea == _cLineaRRGG)
                {
                    //Busca datos de Póliza por Nro. de Póliza:
                    _cElemento = "Buscar póliza";
                    _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_nro).Value);

                    //Obtener nombre de la oferta:
                    _cNombreOferta = _Funciones.FindElement(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_PolicyDV:Offering"), Convert.ToInt32(_TiempoEspera[3])).Text;

                    IniciarCambioPoliza(oTicketDatos);
                    if (!FormularioCambioPoliza(oTicketDatos)) { _bControl = true; return; }
                    if (!String.IsNullOrWhiteSpace(_cNombreOferta)) { SeleccionarOferta(oTicketDatos); }

                    //Método para actualizar datos:
                    GetFieldIdByNames(oTicketDatos, eesFields.Default.listacampos);

                    if (AnalisisDeRiesgos())
                    {
                        ConfirmarTrabajo();
                    }
                    else
                    {
                        //Cancelar cotización:
                        _cElemento = "Cancelar cotización";
                        _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:JobWizardToolbarButtonSet:WithdrawJob")).Click();
                        _Funciones.VerificarVentanaAlerta(_driverGlobal);
                        _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[2]));
                    }
                }
                else if (_cLinea == _cLineaAlianzas)
                {
                    //Busca datos de Póliza por Nro. de Póliza:
                    _cElemento = "Buscar póliza";
                    _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_nro).Value);

                    //Obtener nombre de la oferta:
                    _cNombreOferta = _Funciones.FindElement(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_PolicyDV:Offering"), Convert.ToInt32(_TiempoEspera[3])).Text;

                    IniciarCambioPoliza(oTicketDatos);
                    if (!FormularioCambioPoliza(oTicketDatos)) { _bControl = true; return; }
                    _Funciones.VerificarVentanaAlerta(_driverGlobal);
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[2]));

                    if (!String.IsNullOrWhiteSpace(_cNombreOferta)) { SeleccionarOferta(oTicketDatos); }

                    //Método para actualizar datos:
                    GetFieldIdByNames(oTicketDatos, eesFields.Default.listacampos);

                    if (AnalisisDeRiesgos())
                    {
                        ConfirmarTrabajo();
                    }
                    else
                    {
                        //Cancelar cotización:
                        _cElemento = "Cancelar cotización";
                        _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:JobWizardToolbarButtonSet:WithdrawJob")).Click();
                        _Funciones.VerificarVentanaAlerta(_driverGlobal);
                        _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[2]));
                    }
                }
                else if (_cLinea == _cLineaLLPP)
                {
                    _nFormulario = 3;
                    IniciarCambioPoliza(oTicketDatos, false);
                    //Método para actualizar datos:
                    GetFieldIdByNames(oTicketDatos, eesFields.Default.listacampos);
                }
            }
            catch (Exception Ex) { LogFailStep(12, Ex); throw new Exception(Ex.Message + " :" + _cElemento, Ex); }
        }

        //Inicia acción Cambiar Póliza:
        private void IniciarCambioPoliza(Ticket oTicketDatos, bool bPoliza = true)
        {
            try
            {
                if (bPoliza)
                {
                    //Ingresar mediante Menú Acciones:
                    _cElemento = "Menú Acciones";
                    _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions")).Click();

                    //Hacer clic en Cambiar Póliza:
                    _cElemento = "Opción Cambiar Póliza";
                    _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions:PolicyFileMenuActions_NewWorkOrder:PolicyFileMenuActions_ChangePolicy")).Click();
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                }
                else
                {
                    //Buscar Documento:
                    _cElemento = "Buscar por documento";
                    _Funciones.BuscarDocumentoPolicyCenter(_driverGlobal,
                        oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nro_dni).Value.Length == 0 ?
                        oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nro_ruc).Value :
                        oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nro_dni).Value);

                    if (_Funciones.ExisteElemento(_driverGlobal, By.Id("ContactFile_AccountsSearch:AssociatedAccountsLV"), _nIntentosPolicyCenter))
                    {
                        //Obtener número de filas de la tabla:
                        _cElemento = "Tabla de cuentas registradas";
                        int nFilas = _Funciones.ObtenerFilasTablaHTML(_driverGlobal, "ContactFile_AccountsSearch:AssociatedAccountsLV");

                        if (nFilas > 1)
                        {
                            for (int i = _nIndexFilaCuenta; i < nFilas; i++)
                            {
                                _cElemento = "Obteniendo Nro. de Cuenta";
                                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.cuenta_nro).Value == _Funciones.ObtenerValorCeldaTabla(_driverGlobal, "ContactFile_AccountsSearch:AssociatedAccountsLV", i, _nIndexCeldaCuenta))
                                {
                                    _cElemento = "Clic en Nro. de Cuenta";
                                    _driverGlobal.FindElement(By.XPath("id('ContactFile_AccountsSearch:AssociatedAccountsLV')/tbody/tr[" + i + "]/td[" + _nIndexCeldaCuenta + "]")).Click();
                                }
                            }
                        }
                        else
                        {
                            //Clic en Nro. de Cuenta:
                            _cElemento = "Clic en Nro. de Cuenta";
                            _driverGlobal.FindElement(By.Id("ContactFile_AccountsSearch:AssociatedAccountsLV:0:AccountNumber")).Click();
                            _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[1]));
                        }
                    }
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al iniciar el cambio de la póliza: " + Ex.Message + " " + _cElemento, Ex); }
        }

        //Formulario inicial para el Cambio de Póliza:
        private Boolean FormularioCambioPoliza(Ticket oTicketDatos)
        {
            try
            {
                //Formulario registro de endoso:
                _cElemento = "Fecha efectiva del cambio";
                //Validar fecha:
                if (ValidarFechaSolicitud(oTicketDatos, Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_efectiva).Value)))
                {
                    _Funciones.FindElement(_driverGlobal, By.XPath("//*[@id='StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:0']/tbody/tr[1]/td[5]/input[@class='textBox']"), _nIntentosPolicyCenter).SendKeys("");
                    _driverGlobal.FindElement(By.XPath("//*[@id='StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:0']/tbody/tr[1]/td[5]/input[@class='textBox']")).SendKeys(Keys.Control + "e");
                    _driverGlobal.FindElement(By.XPath("//*[@id='StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:0']/tbody/tr[1]/td[5]/input[@class='textBox']")).
                        SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_efectiva).Value.Replace("/", ""));
                }

                else { return false; }

                if (_Funciones.ExisteElemento(_driverGlobal, By.Id("StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:TypeReason"), _nIntentosPolicyCenter))
                {
                    //Seleccionar tipo de complejidad:
                    _cElemento = "Tipo de Complejidad";
                    string cComplejidad = _Funciones.GetDomainValue(Convert.ToInt32(_DominioComplejidad[0]), Convert.ToInt32(_DominioComplejidad[1]), Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.idcomplejidad).Value)).ToUpperInvariant();
                    _Funciones.SeleccionarCombo(_driverGlobal, "StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:TypeReason", cComplejidad);
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                }

                //Seleccionar motivo del endoso:
                _cElemento = "Motivo del endoso";
                string cMotivoEndoso = _Funciones.GetDomainValue(Convert.ToInt32(_DominioMotivo[0]), Convert.ToInt32(_DominioMotivo[1]), Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.idmotivoendoso).Value)).ToUpperInvariant();
                _Funciones.SeleccionarCombo(_driverGlobal, "StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:Description", cMotivoEndoso);
                _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));

                //Ingresar comentarios adicionales:
                _cElemento = "Comentarios adicionales";
                _driverGlobal.FindElement(By.Id("StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:Comments")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.comentarios_adicionales).Value);

                //Clic en Siguiente:
                _cElemento = "Botón Siguiente";
                _Funciones.FindElement(_driverGlobal, By.Id("StartPolicyChange:StartPolicyChangeScreen:NewPolicyChange")).Click();
                _Funciones.VerificarVentanaAlerta(_driverGlobal);
                _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[1]));
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error en formulario: " + Ex.Message + " " + _cElemento, Ex); }
            return true;
        }

        //Verifica si se requiere seleccionar Oferta:
        private void SeleccionarOferta(Ticket oTicketDatos)
        {
            try
            {
                if (_Funciones.ExisteElemento(_driverGlobal, By.Id("PolicyChangeWizard:OfferingScreen:OfferingSelection"), _nIntentosPolicyCenter))
                {
                    if (_cLinea != _cLineaRRGG)
                    {
                        //Seleccionar oferta:
                        _cElemento = "Seleccionar oferta";
                        _Funciones.SeleccionarCombo(_driverGlobal, "PolicyChangeWizard:OfferingScreen:OfferingSelection", _cNombreOferta.ToUpperInvariant());
                    }

                    //Clic en Siguiente:
                    _cElemento = "Botón siguiente";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Next")).Click();
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[1]));
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al seleccionar oferta: " + Ex.Message + " " + _cElemento, Ex); }
        }

        //Valida si fecha de solicitud está dentro del rango de vigencia:
        private Boolean ValidarFechaSolicitud(Ticket oTicketDatos, DateTime dFecha)
        {
            if (dFecha < Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_fec_ini_vig).Value)) { return false; }
            else if (dFecha >= Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_fec_fin_vig).Value)) { return false; }
            return true;
        }

        private void FormularioEditarPoliza(int nFieldId)
        {
            try
            {
                //Clic en Nombre del asegurado:
                _cElemento = "Nombre del Asegurado";
                _Funciones.FindElement(_driverGlobal, By.Id("PolicyChangeWizard:LOBWizardStepGroup:PolicyChangeWizard_PolicyInfoScreen:PolicyChangeWizard_PolicyInfoDV:AccountInfoInputSet:Name"), Convert.ToInt32(_TiempoEspera[1])).Click();

                //Verifica si es Persona o Empresa:
                if (_valoresTickets[0].Length > 0)
                {
                    //Es persona:
                    if (nFieldId == eesFields.Default.nombre_s)
                    {
                        //Nombre:
                        _cElemento = "Nombre";
                        _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:FirstName"));
                        _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:FirstName")).SendKeys(_cTicketValue);
                    }
                    else if (nFieldId == eesFields.Default.apellido_paterno)
                    {
                        //Apellido Paterno:
                        _cElemento = "Apellido Paterno";
                        _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:LastName"));
                        _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:LastName")).SendKeys(_cTicketValue);
                    }
                    else if (nFieldId == eesFields.Default.apellido_materno)
                    {
                        //Apellido Materno:
                        _cElemento = "Apellido Materno";
                        _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:LastName2"));
                        _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:LastName2")).SendKeys(_cTicketValue);
                    }
                    else if (nFieldId == eesFields.Default.fecha_de_nacimiento)
                    {
                        //Fecha de nacimiento:
                        _cElemento = "Fecha de nacimiento";
                        _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:DateOfBirth"));
                        _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:DateOfBirth")).SendKeys(_cTicketValue);
                    }
                    else if (nFieldId == eesFields.Default.sexo)
                    {
                        //Sexo:
                        _cElemento = "Sexo";
                        _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:Gender"));
                        _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:Gender")).SendKeys(_cTicketValue);
                    }
                    else if (nFieldId == eesFields.Default.estado_civil)
                    {
                        //Estado civil:
                        _cElemento = "Estado civil";
                        _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:MaritalStatus"));
                        _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:MaritalStatus")).SendKeys(_cTicketValue);
                    }
                    else if (nFieldId == eesFields.Default.pais_del_telefono)
                    {
                        //Código País del teléfono:
                        _cElemento = "Código país";
                        _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:PhoneContactInputSet:CountryHomeTelephone", _cTicketValue.ToUpperInvariant());
                        _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                    }
                    else if (nFieldId == eesFields.Default.indicativo_codigo_de_area)
                    {
                        //Código Ciudad:
                        _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:PhoneContactInputSet:HomePhoneCityCodeExt2", _cTicketValue.ToUpperInvariant());
                        _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                    }
                    else if (nFieldId == eesFields.Default.telefono_de_casa)
                    {
                        //Teléfono de domicilio:
                        _cElemento = "Teléfono de domicilio";
                        _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:PhoneContactInputSet:HomePhone"));
                        _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:PhoneContactInputSet:HomePhone")).SendKeys(_cTicketValue);
                    }
                    else if (nFieldId == eesFields.Default.correo_personal)
                    {
                        //Correo principal:
                        _cElemento = "Seleccionar correo principal";
                        _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:ContactEmailsInputSet:PrimaryEmailTypeExt", "PERSONAL");
                        _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[1]));

                        _cElemento = "Correo electrónico personal";
                        _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:ContactEmailsInputSet:PrimaryEmailTypeExt"));
                        _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:ContactEmailsInputSet:PrimaryEmailTypeExt")).SendKeys(_cTicketValue.ToUpperInvariant());
                    }
                }
                else if (_valoresTickets[1].Length > 0)
                {
                    //Es empresa:
                    if (nFieldId == eesFields.Default.razon_social)
                    {
                        //Razón social:
                        _cElemento = "Razón Social";
                        _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:CompanyName"));
                        _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:CompanyName")).SendKeys(_cTicketValue);
                    }
                    else if (nFieldId == eesFields.Default.nombre_comercial)
                    {
                        //Nombre comercial:
                        _cElemento = "Nombre comercial";
                        _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:CommercialNameExt"));
                        _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:CommercialNameExt")).SendKeys(_cTicketValue);
                    }
                    else if (nFieldId == eesFields.Default.sector_economico)
                    {
                        //Sector económico:
                        _cElemento = "Sector económico";
                        _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:EconomicSector", _cTicketValue.ToUpperInvariant());
                        _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[1]));
                    }
                    else if (nFieldId == eesFields.Default.actividad_economica)
                    {
                        //Actividad económica:
                        _cElemento = "Actividad económica";
                        _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:EconomicActivity", _cTicketValue.ToUpperInvariant());
                    }
                }

                //Campos comunes:
                if (nFieldId == eesFields.Default.pais_de_procedencia)
                {
                    //País de procedencia:
                    _cElemento = "País de procedencia";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:CountryOfOrigin", _cTicketValue.ToUpperInvariant());
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[1]));
                }
                else if (nFieldId == eesFields.Default.tipo_de_direccion)
                {
                    //Tipo de dirección:
                    _cElemento = "Tipo de dirección";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressType", _cTicketValue.ToUpperInvariant());
                }
                else if (nFieldId == eesFields.Default.pais)
                {
                    //País:
                    _cElemento = "País";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_Country", _cTicketValue.ToUpperInvariant());
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                }
                else if (nFieldId == eesFields.Default.departamento)
                {
                    //Departamento:
                    _cElemento = "Departamento";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_Department", _cTicketValue.ToUpperInvariant());
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                }
                else if (nFieldId == eesFields.Default.provincia)
                {
                    //Provincia:
                    _cElemento = "Provincia";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_Province", _cTicketValue);
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                }
                else if (nFieldId == eesFields.Default.distrito)
                {
                    //Distrito:
                    _cElemento = "Distrito";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_District", _cTicketValue);
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                }
                else if (nFieldId == eesFields.Default.tipo_de_calle)
                {
                    //Tipo de calle:
                    _cElemento = "Tipo de calle";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_StreetType", _cTicketValue);
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                }
                else if (nFieldId == eesFields.Default.nombre_de_la_calle)
                {
                    //Nombre de la calle:
                    _cElemento = "Nombre de la calle";
                    _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressLine1"));
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressLine1")).SendKeys(_cTicketValue);
                }
                else if (nFieldId == eesFields.Default.numero)
                {
                    //Número:
                    _cElemento = "Número";
                    _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressLine2"));
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressLine2")).SendKeys(_cTicketValue);
                }
                else if (nFieldId == eesFields.Default.referencia)
                {
                    //Referencia:
                    _cElemento = "Referencia";
                    _Funciones.LimpiarElementoInput(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressLine3"));
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressLine3")).SendKeys(_cTicketValue);
                }

                //Finalizar formulario:
                //Clic en botón Aceptar:
                _cElemento = "Botón Aceptar";
                _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:Update")).Click();

                if (_Funciones.ExisteElemento(_driverGlobal, By.Id("EditPolicyContactRolePopup:ContactDetailScreen:_msgs_msgs")))
                {
                    //Clic en Cancelar:
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:Cancel")).Click();
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[1]));
                    FormularioEditarPoliza(_nFieldId);
                }
                else
                {
                    //Clic en Cotización:
                    _cElemento = "Botón Cotización";
                    _Funciones.FindElement(_driverGlobal, By.Id("PolicyChangeWizard:LOBWizardStepGroup:PolicyChangeWizard_PolicyInfoScreen:JobWizardToolbarButtonSet:QuoteOrReview"), Convert.ToInt32(_TiempoEspera[1])).Click();

                    if (_Funciones.ExisteElemento(_driverGlobal, By.Id("UWBlockProgressIssuesPopup:IssuesScreen:DetailsButton"), _nIntentosPolicyCenter))
                    {
                        //Clic en botón Detalles:
                        _cElemento = "Botón Detalles";
                        _driverGlobal.FindElement(By.Id("UWBlockProgressIssuesPopup:IssuesScreen:DetailsButton")).Click();
                        _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                    }
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error en formulario Editar Cuenta: " + Ex.Message + " " + _cElemento, Ex); }
        }

        private void FormularioEditarCuenta(int nFieldId)
        {
            try
            {
                //Clic en Editar Cuenta:
                _cElemento = "Editar cuenta";
                _driverGlobal.FindElement(By.Id("AccountFile_Summary:AccountFile_SummaryScreen:EditAccount")).Click();
                _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[1]));

                if (nFieldId == eesFields.Default.nombre_s)
                {
                    //Nombre:
                    _cElemento = "Nombre";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:FirstName")).SendKeys(_cTicketValue);
                }
                else if (nFieldId == eesFields.Default.apellido_paterno)
                {
                    //Apellido Paterno:
                    _cElemento = "Apellido paterno";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:LastName")).SendKeys(_cTicketValue);
                }
                else if (nFieldId == eesFields.Default.apellido_materno)
                {
                    //Apellido Materno:
                    _cElemento = "Apellido materno";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:SecondLastNameExt")).SendKeys(_cTicketValue);
                }
                else if (nFieldId == eesFields.Default.pais_de_procedencia)
                {
                    //País de procedencia:
                    _cElemento = "País de procedencia";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:ForeignCountry", _cTicketValue);
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                }
                else if (nFieldId == eesFields.Default.fecha_de_nacimiento)
                {
                    //Fecha de nacimiento:
                    _cElemento = "Fecha de nacimiento";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:DateOfBirth")).SendKeys(_cTicketValue);
                }
                else if (nFieldId == eesFields.Default.tipo_de_direccion)
                {
                    //Tipo de dirección:
                    _cElemento = "Tipo de dirección";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_AddressType", _cTicketValue);
                }
                else if (nFieldId == eesFields.Default.pais)
                {
                    //País:
                    _cElemento = "País";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_Country", _cTicketValue);
                }
                else if (nFieldId == eesFields.Default.departamento)
                {
                    //Departamento:
                    _cElemento = "Departamento";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_Department", _cTicketValue);
                }
                else if (_nFieldId == eesFields.Default.provincia)
                {
                    //Provincia:
                    _cElemento = "Provincia";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_Province", _cTicketValue);
                }
                else if (_nFieldId == eesFields.Default.distrito)
                {
                    //Distrito:
                    _cElemento = "Distrito";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_District", _cTicketValue);
                }
                else if (_nFieldId == eesFields.Default.tipo_de_calle)
                {
                    //Tipo de calle:
                    _cElemento = "Tipo de calle";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_StreetType", _cTicketValue);
                }
                else if (_nFieldId == eesFields.Default.nombre_de_la_calle)
                {
                    //Nombre de la calle:
                    _cElemento = "Nombre de la calle";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_AddressLine1")).SendKeys(_cTicketValue);
                }
                else if (_nFieldId == eesFields.Default.numero)
                {
                    //Número:
                    _cElemento = "Número";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_AddressLine2")).SendKeys(_cTicketValue);
                }

                //Clic en Actualizar:
                _cElemento = "Clic en Actualizar";
                _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:Update")).Click();
                _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[1]));

                //Verificar si los datos son correctos:
                if (_Funciones.ExisteElemento(_driverGlobal, By.Id("EditAccountPopup:EditAccountScreen:_msgs_msgs"), _nIntentosPolicyCenter))
                {
                    //Clic en Cancelar:
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:Cancel")).Click();
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[1]));
                    FormularioEditarCuenta(_nFieldId);
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error en formulario Editar Cuenta: " + Ex.Message + " " + _cElemento, Ex); }
        }

        //Pasos en Ventana Análisis de Riesgo:
        private Boolean AnalisisDeRiesgos()
        {
            int nFilas = 0;
            try
            {
                //Ventana de aprobación de bloqueantes:
                if (_Funciones.ExisteElemento(_driverGlobal, By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:1")))
                {
                    //Verifica si existen bloqueantes:
                    _cElemento = "Tabla de bloqueantes";
                    nFilas = _Funciones.ObtenerFilasTablaHTML(_driverGlobal, "PolicyChangeWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:1", _cCeldaLimitante);
                }

                if (nFilas > 1)
                {
                    for (int i = 1; i < nFilas; i++)
                    {
                        //Marcar check de aprobación:
                        _cElemento = "Check bloqueante de cotización";
                        _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:" + i + ":UWIssueRowSet:_Checkbox")).Click();
                    }

                    if (_Funciones.ExisteElemento(_driverGlobal, By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:Approve")))
                    {
                        //Clic en botón Aprobar:
                        _cElemento = "Botón Aprobar";
                        _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:Approve")).Click();
                        _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));

                        //Clic en opción SI:
                        _cElemento = "Permitir edición";
                        _driverGlobal.FindElement(By.Id("RiskApprovalDetailsPopup:0:IssueDetailsDV:UWApprovalLV:EditBeforeBind_true")).Click();
                        _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));

                        //Clic en Botón Aceptar:
                        _cElemento = "Botón Aceptar";
                        _driverGlobal.FindElement(By.Id("RiskApprovalDetailsPopup:Update")).Click();
                        _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));
                    }
                    else { return false; }
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error en análisis de riesgos: " + Ex.Message + " " + _cElemento, Ex); }
            return true;
        }

        private void ConfirmarTrabajo()
        {
            try
            {
                if (_Funciones.ExisteElemento(_driverGlobal, By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:JobWizardToolbarButtonSet:QuoteOrReview"), _nIntentosPolicyCenter))
                {
                    //Clic en Cotización:
                    _cElemento = "Botón Cotización";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:JobWizardToolbarButtonSet:QuoteOrReview")).Click();
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[2]));
                }

                if (_Funciones.ExisteElemento(_driverGlobal, By.Id("PolicyChangeWizard:PolicyChangeWizard_QuoteScreen:RatingCumulDetailsPanelSet:RatingOverrideButtonDV:RatingOverrideButtonDV:OverrideRating_link"), _nIntentosPolicyCenter))
                {
                    //Clic en Reescribir prima:
                    _cElemento = "Reescribir prima y comisiones";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:PolicyChangeWizard_QuoteScreen:RatingCumulDetailsPanelSet:RatingOverrideButtonDV:RatingOverrideButtonDV:OverrideRating_link")).Click();
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[0]));

                    //Clic en Recalcular:
                    _cElemento = "Recalcular prima";
                    _driverGlobal.FindElement(By.Id("RatingOverridePopup:Update")).Click();
                    _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[2]));
                }

                //Clic en Confirmar trabajo:
                _cElemento = "Botón Confirmar trabajo";
                _driverGlobal.FindElement(By.Id("PolicyChangeWizard:PolicyChangeWizard_QuoteScreen:JobWizardToolbarButtonSet:BindPolicyChange")).Click();
                _Funciones.VerificarVentanaAlerta(_driverGlobal);
                _Funciones.Esperar(Convert.ToInt32(_TiempoEspera[2]));

                if (_Funciones.ExisteElemento(_driverGlobal, By.Id("JobComplete:JobCompleteScreen:Message")))
                {
                    _cOrdenTrabajo = _Funciones.ObtenerCadenaDeNumeros(_driverGlobal.FindElement(By.Id("JobComplete:JobCompleteScreen:Message")).Text);
                }

                //Cambio de póliza Completada:
                _cElemento = "Cambio completado";
                _driverGlobal.FindElement(By.Id("JobComplete:JobCompleteScreen:JobCompleteDV:ViewPolicy")).Click();
                _Funciones.Esperar();
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al confirmar trabajo: " + Ex.Message + " " + _cElemento, Ex); }
        }

        private void AgregarValoresTicket(Ticket oTicketDatos)
        {
            try
            {
                oTicketDatos.TicketValues.Add(new TicketValue{ ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.endoso_nro, Value = _cOrdenTrabajo });
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al agregar valores al ticket " + Convert.ToString(oTicketDatos.Id) + ": " + Ex.Message + " " + _cElemento, Ex); }
        }
        #endregion
    }
}
