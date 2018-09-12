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
        private static string Acceso = string.Empty;
        static string[] _valoresTickets = new string[10];
        static string[] _valoresTickets_Ident = new string[10];
        private string _urlContactManager = string.Empty;
        private string _usuarioContactManager = string.Empty;
        private string _contraseñaContactManager = string.Empty;
        private static int _nIdEstadoError; //Mesa de Control
        private static int _nIdEstadoSiguiente; //Robot Crear Ticket Hijo
        private static int _nIntentosPolicyCenter;
        private static string _cUrlPolicyCenter = string.Empty;
        private static string _cComentariosAdicionales = string.Empty;
        private static string[] _Usuarios;
        private static int _nIndice;
        private static string _cElemento = string.Empty;
        private static string _cLineaAutos = string.Empty;
        private static string _cLineaRRGG = string.Empty;
        private static string _cLineaAlianzas = string.Empty;
        private static string _cLineaLLPP = string.Empty;
        private static string _cLinea = string.Empty;
        List<TicketValue> ticketValue = null;
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

            ObtenerParametros();
            LogStartStep(4);
            foreach (Ticket oTicket in _oRobot.Tickets)
            {
                try
                {
                    _oMesaControl = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdEstadoError);
                    _oTicketHijo = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdEstadoSiguiente);
                    //Obteniendo Línea de Negocio:
                    _cLinea = _Funciones.ObtenerValorDominio(oTicket, Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_linea).Value));
                    ProcesarTicket(oTicket);
                }
                catch (Exception Ex) { CambiarEstadoTicket(oTicket, _oMesaControl, Ex.Message); LogFailStep(30, Ex); }
                finally { _Funciones.CerrarDriver(_driverGlobal); }
            }
        }

        //Obtiene los parámetros asociados al robot:
        private void ObtenerParametros()
        {
            try
            {
                _urlContactManager = _oRobot.GetValueParamRobot("URLContactManager").ValueParam;
                _usuarioContactManager = _oRobot.GetValueParamRobot("UsuarioContactManager").ValueParam;
                _contraseñaContactManager = _oRobot.GetValueParamRobot("PasswordContactManager").ValueParam;
                _nIdEstadoError = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoError").ValueParam);
                _nIdEstadoSiguiente = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoSiguiente").ValueParam);
                _nIntentosPolicyCenter = Convert.ToInt32(_oRobot.GetValueParamRobot("nIntentosPolicyCenter").ValueParam);
                _cUrlPolicyCenter = _oRobot.GetValueParamRobot("URLPolicyCenter").ValueParam;
                _cLineaAutos = _oRobot.GetValueParamRobot("LineaAutos").ValueParam;
                _cLineaRRGG = _oRobot.GetValueParamRobot("LineaRRGG").ValueParam;
                _cLineaAlianzas = _oRobot.GetValueParamRobot("LineaAlianzas").ValueParam;
                _cLineaLLPP = _oRobot.GetValueParamRobot("LineaLLPP").ValueParam;
                _cComentariosAdicionales = _oRobot.GetValueParamRobot("ComentariosAdicionales").ValueParam;
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

        #region CONTACT MANAGER
        private void ContactManager(Ticket oTicket)
        {
            _Funciones.AbrirSelenium(ref _driverGlobal);

            _Funciones.NavegarUrlPolicyCenter(_driverGlobal, _urlContactManager);
            _Funciones.LoginPolicyCenter(_driverGlobal, _usuarioContactManager, _contraseñaContactManager);
            ActualizarContactManager(oTicket);
        }

        private void ActualizarContactManager(Ticket ticket)
        {
            //Opteniendo DNI
            _valoresTickets[0] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.dni).Value;
            //Opteniendo RUC
            _valoresTickets[1] = ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.ruc).Value;

            String mensajeError_Xpath = "//*[@id='ABContactSearch:ABContactSearchScreen:_msgs_msgs']/div";

            if (_valoresTickets[0] != "" && _valoresTickets[1] != "")
            {
                if (_valoresTickets[0] != "")
                {
                    Filtros(1, _valoresTickets[0]);

                    if (!_Funciones.ExisteElementoXPath(_driverGlobal, mensajeError_Xpath, 1))
                    {
                        AccederRegistro();
                        Acceso = "";
                    }
                    else
                    {
                        Acceso = "Negativo";
                    }
                }
                if (_valoresTickets[1] != "" && Acceso == "Negativo")
                {
                    Filtros(2, _valoresTickets[1]);


                    if (!_Funciones.ExisteElementoXPath(_driverGlobal, mensajeError_Xpath, 1))
                    {
                        AccederRegistro();
                        Acceso = "";
                    }
                    else
                    {
                        Acceso = "Negativo";
                    }
                }
            }
            else
            {
                if (_valoresTickets[0] != "")
                {
                    Filtros(1, _valoresTickets[0]);

                    if (!_Funciones.ExisteElementoXPath(_driverGlobal, mensajeError_Xpath, 1))
                    {
                        AccederRegistro();
                        Acceso = "";
                    }
                    else
                    {
                        Acceso = "Negativo";
                    }
                }

                if (_valoresTickets[1] != "")
                {
                    Filtros(2, _valoresTickets[1]);

                    if (!_Funciones.ExisteElementoXPath(_driverGlobal, mensajeError_Xpath, 1))
                    {
                        AccederRegistro();
                        Acceso = "";
                    }
                    else
                    {
                        Acceso = "Negativo";
                    }
                }
            }

            if (Acceso == "Negativo")
            {
                _oRobot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, "Registro no Encontrado"), _oRobot.GetNextStateAction(ticket).First(o => o.DestinationStateId == Convert.ToInt32(_nIdEstadoSiguiente)).Id);
            }
            else
            {
                ValoresTicketsRobots(ticket);
            }
        }

        public void Filtros(int indicador, string valor)
        {
            if (indicador == 1)
            {
                IngresarTextoSelect("ABContactSearch:ABContactSearchScreen:ContactSearchDV:ContactSubtype", "Persona");
                IngresarTextoSelect("ABContactSearch:ABContactSearchScreen:ContactSearchDV:PrimaryOfficialIDTypeExt", "DNI");
                _Funciones.Esperar(2);
                _driverGlobal.FindElement(By.Id("ABContactSearch:ABContactSearchScreen:ContactSearchDV:TaxID")).SendKeys(Keys.Control + "e");
                _Funciones.Esperar(2);
                _driverGlobal.FindElement(By.Id("ABContactSearch:ABContactSearchScreen:ContactSearchDV:TaxID")).SendKeys(valor);
                _Funciones.Esperar(4);
                _driverGlobal.FindElement(By.ClassName("bigButton_link")).Click();
                _Funciones.Esperar(2);
            }
            else
            {
                IngresarTextoSelect("ABContactSearch:ABContactSearchScreen:ContactSearchDV:ContactSubtype", "Empresa");
                IngresarTextoSelect("ABContactSearch:ABContactSearchScreen:ContactSearchDV:PrimaryOfficialIDTypeExt", "RUC");
                _Funciones.Esperar(2);
                _driverGlobal.FindElement(By.Id("ABContactSearch:ABContactSearchScreen:ContactSearchDV:TaxID")).SendKeys(Keys.Control + "e");
                _Funciones.Esperar(2);
                _driverGlobal.FindElement(By.Id("ABContactSearch:ABContactSearchScreen:ContactSearchDV:TaxID")).SendKeys(valor);
                _Funciones.Esperar(4);
                _driverGlobal.FindElement(By.ClassName("bigButton_link")).Click();
                _Funciones.Esperar(2);
            }
        }

        public void AccederRegistro()
        {
            _driverGlobal.FindElement(By.Id("ABContactSearch:ABContactSearchScreen:ContactSearchResultsLV:0:DisplayName")).Click();
            _Funciones.Esperar(3);
            _driverGlobal.FindElement(By.XPath("//*[@id='ContactDetail:ABContactDetailScreen:ContactBasicsDV_tb:Edit']/span[2]")).Click();
            _Funciones.Esperar(2);
        }

        private void ValoresTicketsRobots(Ticket ticket)
        {
            try
            {

                var container = ODataContextWrapper.GetContainer();

                ticket = _oRobot.Tickets.FirstOrDefault();

                ticketValue = _oRobot.GetDataQueryTicketValue().Where(a => a.TicketId == ticket.Id).ToList();


                String[] listaCampos = (ticket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.listacampos).Value).Split(',');

                foreach (string campo in listaCampos)
                {
                    var Fiel = container.Fields.Where(f => f.Name == campo.ToString()).Select(t => new { t.Id, t.Label }).FirstOrDefault();
                    var texto = ticketValue.Where(t => t.FieldId == Convert.ToInt32(Fiel.Id)).Select(t => new { t.Value }).FirstOrDefault();

                    EditarFormulario(Fiel.Label.ToString(), texto.Value.ToString());
                }

                //Guardar Cambios
                _Funciones.Esperar(3);
                _driverGlobal.FindElement(By.XPath("//*[@id='ContactDetail:ABContactDetailScreen:ContactBasicsDV_tb:Update']/span[2]")).Click();
                _Funciones.Esperar(3);

                _oRobot.SaveTicketNextState(ticket, _nIdEstadoSiguiente);

            }
            catch (Exception ex)
            {

                throw;
            }

        }

        public void EditarFormulario(String campo, String texto)
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
                    IngresarTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_District", texto);
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

        public void EscribirElementoXPathActualizarDatos(int index, int posicion, string texto, string clase)
        {
            //IWebElement _element = null; ;
            string xPath = "//*[@id='ContactDetail:ABContactDetailScreen:ContactBasicsDV:" + index + "']/tbody/tr[" + posicion + "]/td[5]/input[@class='" + clase + "']";
            _driverGlobal.FindElement(By.XPath(xPath)).SendKeys("");
            _Funciones.Esperar(2);
            _driverGlobal.FindElement(By.XPath(xPath)).SendKeys(Keys.Control + "e");
            _driverGlobal.FindElement(By.XPath(xPath)).SendKeys(texto);

            //_element.SendKeys(Keys.Control + "e");
            //_element.SendKeys(Texto);
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
            _Funciones.Esperar(2);
        }
        #endregion

        #region POLICYCENTER
        private void PolicyCenter(Ticket oTicket)
        {
            if (ValidarVacios(oTicket))
            {
                _nIndice = 1;
                for (int i = 0; i < _nIntentosPolicyCenter; i++)
                {
                    _Funciones.AbrirSelenium(ref _driverGlobal);
                    _Funciones.NavegarUrlPolicyCenter(_driverGlobal, _cUrlPolicyCenter);
                    Credenciales();
                    _Funciones.LoginPolicyCenter(_driverGlobal, _Usuarios[0], _Usuarios[1]);
                    if (_Funciones.ExisteElemento(_driverGlobal, "TabBar:PolicyTab_arrow", _nIntentosPolicyCenter))
                    {
                        break;
                    }
                    _Funciones.CerrarDriver(_driverGlobal);
                    _nIndice += 1;
                }
                ActualizarPolicyCenter(oTicket);

                //Si todo es conforme, pasa al estado Crear Ticket Hijo:
                CambiarEstadoTicket(oTicket, _oTicketHijo);
                LogEndStep(4);
            }
            else
            {
                //Enviar a mesa de control:
                CambiarEstadoTicket(oTicket, _oMesaControl, "El ticket " + Convert.ToString(oTicket.Id) + " no cuenta con todos los datos necesarios.");
            }
        }

        //Obtiene los usuarios con sus respectivas contraseñas:
        private string[] Credenciales()
        {
            try
            {
                //Usuario y contraseña de Autos:
                _Usuarios = _oRobot.GetValueParamRobot("AccesoPCyCM_" + _nIndice).ValueParam.Split(',');
                return _Usuarios;
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al obtener datos de usuario: " + Ex.Message, Ex); }
        }

        //Valida que no tenga campos vacíos:
        private Boolean ValidarVacios(Ticket oTicketDatos)
        {
            try
            {
                int[] oCampos = new int[] { eesFields.Default.nombre_contratante, eesFields.Default.asegurado_nombre,
                                            eesFields.Default.email_solicitante, eesFields.Default.fecha_hora_de_email,
                                            eesFields.Default.date_inicio_vigencia, eesFields.Default.date_fin_vigencia};

                return _Funciones.ValidarCamposVacios(oTicketDatos, oCampos);
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al validar campos del Ticket: " + Convert.ToString(oTicketDatos.Id), Ex); }
        }

        private void ActualizarPolicyCenter(Ticket oTicketDatos)
        {
            try
            {
                LogStartStep(1);

                if (_cLinea == _cLineaAutos)
                {
                    //Busca datos de Póliza por Nro. de Póliza:
                    _cElemento = "Buscar póliza";
                    _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_nro).Value);
                    IniciarCambioPoliza(oTicketDatos);
                    FormularioCambioPoliza(oTicketDatos);
                    SeleccionarOferta(oTicketDatos);

                    //Clic en Nombre del asegurado:
                    _cElemento = "Nombre del Asegurado";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:LOBWizardStepGroup:PolicyChangeWizard_PolicyInfoScreen:PolicyChangeWizard_PolicyInfoDV:AccountInfoInputSet:Name")).Click();
                    _Funciones.Esperar(7);

                    //Método para actualizar datos:
                    FormularioEditarCuenta(oTicketDatos);

                    if (AnalisisDeRiesgos())
                    {
                        ConfirmarTrabajo();
                    }
                    else
                    {
                        //Cancelar cotización:
                        _cElemento = "Cancelar cotización";
                        _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:JobWizardToolbarButtonSet:WithdrawJob")).Click();
                        _Funciones.VentanaMensajeWeb(_driverGlobal);
                        _Funciones.Esperar(6);
                    }
                }
                else if (_cLinea == _cLineaRRGG)
                {
                    //Busca datos de Póliza por Nro. de Póliza:
                    _cElemento = "Buscar póliza";
                    _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_nro).Value);
                    IniciarCambioPoliza(oTicketDatos);
                    FormularioCambioPoliza(oTicketDatos);
                    SeleccionarOferta(oTicketDatos);

                    //Clic en Nombre:
                    _cElemento = "Nombre del Asegurado";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:LOBWizardStepGroup:PolicyChangeWizard_PolicyInfoScreen:PolicyChangeWizard_PolicyInfoDV:AccountInfoInputSet:Name")).Click();
                    _Funciones.Esperar(7);

                    //Método para actualizar datos:
                    FormularioEditarCuentaRRGG(oTicketDatos);

                    if (AnalisisDeRiesgos())
                    {
                        ConfirmarTrabajo();
                    }
                    else
                    {
                        //Cancelar cotización:
                        _cElemento = "Cancelar cotización";
                        _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:JobWizardToolbarButtonSet:WithdrawJob")).Click();
                        _Funciones.VentanaMensajeWeb(_driverGlobal);
                        _Funciones.Esperar(6);
                    }
                }
                else if (_cLinea == _cLineaAlianzas)
                {
                    //Busca datos de Póliza por Nro. de Póliza:
                    _cElemento = "Buscar póliza";
                    _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_nro).Value);
                    IniciarCambioPoliza(oTicketDatos);
                    FormularioCambioPoliza(oTicketDatos);
                    SeleccionarOferta(oTicketDatos);

                    //Clic en Nombre del asegurado:
                    _cElemento = "Nombre del Asegurado";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:LOBWizardStepGroup:PolicyChangeWizard_PolicyInfoScreen:PolicyChangeWizard_PolicyInfoDV:AccountInfoInputSet:Name")).Click();
                    _Funciones.Esperar(8);

                    FormularioEditarCuenta(oTicketDatos);

                    if (AnalisisDeRiesgos())
                    {
                        ConfirmarTrabajo();
                    }
                    else
                    {
                        //Cancelar cotización:
                        _cElemento = "Cancelar cotización";
                        _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:JobWizardToolbarButtonSet:WithdrawJob")).Click();
                        _Funciones.VentanaMensajeWeb(_driverGlobal);
                        _Funciones.Esperar(6);
                    }
                }
                else if (_cLinea == _cLineaLLPP)
                {
                    IniciarCambioPoliza(oTicketDatos, false);
                    FormularioEditarCuenta(oTicketDatos);

                    //Verificar si los datos son correctos:
                    if (_Funciones.ExisteElemento(_driverGlobal, "EditAccountPopup:EditAccountScreen:_msgs_msgs", 2)) { FormularioEditarCuenta(oTicketDatos); }
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
                    _Funciones.Esperar(3);
                }
                else
                {
                    //Buscar Documento:
                    _cElemento = "Buscar por documento";
                    _Funciones.BuscarDocumentoPolicyCenter(_driverGlobal,
                        oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.dni).Value.Length == 0 ?
                        oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.ruc).Value :
                        oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.dni).Value);

                    //Clic en Nro. de Cuenta:
                    _cElemento = "Clic en Nro. de Cuenta";
                    _driverGlobal.FindElement(By.Id("ContactFile_AccountsSearch:AssociatedAccountsLV:0:AccountNumber")).Click();
                    _Funciones.Esperar(5);

                    //Clic en Editar Cuenta:
                    _cElemento = "Editar cuenta";
                    _driverGlobal.FindElement(By.Id("AccountFile_Summary:AccountFile_SummaryScreen:EditAccount")).Click();
                    _Funciones.Esperar(4);
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al iniciar el cambio de la póliza: " + Ex.Message + " " + _cElemento, Ex); }
        }

        //Formulario inicial para el Cambio de Póliza:
        private void FormularioCambioPoliza(Ticket oTicketDatos)
        {
            try
            {
                //Formulario registro de endoso:
                _cElemento = "Fecha efectiva del cambio";
                //Validar fecha:
                if (ValidarFechaSolicitud(oTicketDatos, Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_efectiva).Value)))
                    _driverGlobal.FindElement(By.Id("StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:EffectiveDate")).
                        SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_efectiva).Value);

                //Seleccionar tipo de complejidad:
                _cElemento = "Tipo de Complejidad";
                _Funciones.SeleccionarCombo(_driverGlobal, "StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:TypeReason", _Funciones.ObtenerValorDominio(oTicketDatos, Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.producto_tipo).Value))); //Campo tipo de complejidad
                _Funciones.Esperar(2);

                //Seleccionar motivo del endoso:
                _cElemento = "Motivo del endoso";
                _Funciones.SeleccionarCombo(_driverGlobal, "StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:Description", _Funciones.ObtenerValorDominio(oTicketDatos, Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_anu_motivo).Value)));
                _Funciones.Esperar(2);

                //Ingresar comentarios adicionales:
                _cElemento = "Comentarios adicionales";
                _driverGlobal.FindElement(By.Id("StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:Comments")).SendKeys(_cComentariosAdicionales);

                //Clic en Siguiente:
                _cElemento = "Botón Siguiente";
                _driverGlobal.FindElement(By.Id("StartPolicyChange:StartPolicyChangeScreen:NewPolicyChange")).Click();
                _Funciones.Esperar(5);
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error en formulario: " + Ex.Message + " " + _cElemento, Ex); }
        }

        //Verifica si se requiere seleccionar Oferta:
        private void SeleccionarOferta(Ticket oTicketDatos)
        {
            try
            {
                if (_Funciones.ExisteElemento(_driverGlobal, "PolicyChangeWizard:OfferingScreen:OfferingSelection", 2))
                {
                    //Seleccionar oferta:
                    _cElemento = "Seleccionar oferta";
                    _Funciones.SeleccionarCombo(_driverGlobal, "PolicyChangeWizard:OfferingScreen:OfferingSelection", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == 3).Value);

                    //Clic en Siguiente:
                    _cElemento = "Botón siguiente";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Next")).Click();
                    _Funciones.Esperar(3);
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al seleccionar oferta: " + Ex.Message + " " + _cElemento, Ex); }
        }

        //Valida si fecha de solicitud está dentro del rango de vigencia:
        private Boolean ValidarFechaSolicitud(Ticket oTicketDatos, DateTime dFecha)
        {
            if (dFecha < Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.date_inicio_vigencia).Value))
            {
                return false;
            }
            else if (dFecha >= Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.date_fin_vigencia).Value))
            {
                return false;
            }
            return true;
        }

        private void FormularioEditarCuentaPersona(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre).Value.Length > 0)
                {
                    //Nombre:
                    _cElemento = "Nombre";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:FirstName")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.apellido_paterno).Value.Length > 0)
                {
                    //Apellido Paterno:
                    _cElemento = "Apellidos paterno";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:LastName")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.apellido_paterno).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.apellido_materno).Value.Length > 0)
                {
                    //Apellido Materno:
                    _cElemento = "Apellidos paterno";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:SecondLastNameExt")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.apellido_materno).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais_de_procedencia).Value.Length > 0)
                {
                    //País de procedencia:
                    _cElemento = "País de procedencia";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:ForeignCountry", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais_de_procedencia).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_de_nacimiento).Value.Length > 0)
                {
                    //Fecha de nacimiento:
                    _cElemento = "Fecha de nacimiento";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:DateOfBirth")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_de_nacimiento).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_direccion).Value.Length > 0)
                {
                    //Tipo de dirección:
                    _cElemento = "Tipo de dirección";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_AddressType", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_direccion).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais).Value.Length > 0)
                {
                    //País:
                    _cElemento = "País";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_Country", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.departamento).Value.Length > 0)
                {
                    //Departamento:
                    _cElemento = "Departamento";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_Department", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.departamento).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.provincia).Value.Length > 0)
                {
                    //Provincia:
                    _cElemento = "Provincia";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_Province", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.provincia).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.distrito).Value.Length > 0)
                {
                    //Distrito:
                    _cElemento = "Distrito";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_District", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.distrito).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_calle).Value.Length > 0)
                {
                    //Tipo de calle:
                    _cElemento = "Tipo de calle";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_StreetType", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_calle).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre_de_la_calle).Value.Length > 0)
                {
                    //Nombre de la calle:
                    _cElemento = "Nombre de la calle";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_AddressLine1")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre_de_la_calle).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.numero).Value.Length > 0)
                {
                    //Número:
                    _cElemento = "Número";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_AddressLine2")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.numero).Value);
                }

                if (_cLinea == _cLineaLLPP)
                {
                    //Clic en Actualizar:
                    _cElemento = "Clic en Actualizar";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:Update")).Click();
                    _Funciones.Esperar(4);
                }
                else
                {
                    //Finalizar formulario:
                    //Clic en botón Aceptar:
                    _cElemento = "Botón Aceptar";
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:Update")).Click();
                    _Funciones.Esperar(3);

                    //Clic en Cotización:
                    _cElemento = "Botón Cotización";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:LOBWizardStepGroup:PolicyChangeWizard_PolicyInfoScreen:JobWizardToolbarButtonSet:QuoteOrReview")).Click();
                    _Funciones.Esperar(10);

                    //Clic en botón Detalles:
                    _cElemento = "Botón Detalles";
                    _driverGlobal.FindElement(By.Id("UWBlockProgressIssuesPopup:IssuesScreen:DetailsButton")).Click();
                    _Funciones.Esperar(3);
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error en formulario Editar Cuenta: " + Ex.Message + " " + _cElemento, Ex); }
        }

        private void FormularioEditarCuentaEmpresa(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais_de_procedencia).Value.Length > 0)
                {
                    //País de procedencia:
                    _cElemento = "País de procedencia";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:CountryOfOrigin", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais_de_procedencia).Value);
                    _Funciones.Esperar(4);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.razon_social).Value.Length > 0)
                {
                    //Razón social:
                    _cElemento = "Razón Social";
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:CompanyName")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.razon_social).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.sector_economico).Value.Length > 0)
                {
                    //Sector económico:
                    _cElemento = "Sector económico";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:EconomicSector", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.sector_economico).Value);
                    _Funciones.Esperar(4);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.actividad_economica).Value.Length > 0)
                {
                    //Actividad económica:
                    _cElemento = "Actividad económica";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:EconomicActivity", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.actividad_economica).Value);
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error en formulario Editar Cuenta: " + Ex.Message + " " + _cElemento, Ex); }
        }

        private void FormularioEditarCuentaRRGG(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre).Value.Length > 0)
                {
                    //Nombre:
                    _cElemento = "Nombre";
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:FirstName")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.apellido_paterno).Value.Length > 0)
                {
                    //Apellido Paterno:
                    _cElemento = "Apellidos paterno";
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:LastName")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.apellido_paterno).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.apellido_materno).Value.Length > 0)
                {
                    //Apellido Materno:
                    _cElemento = "Apellidos paterno";
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:LastName2")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.apellido_materno).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais_de_procedencia).Value.Length > 0)
                {
                    //País de procedencia:
                    _cElemento = "País de procedencia";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:CountryOfOrigin", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais_de_procedencia).Value);
                    _Funciones.Esperar();
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_de_nacimiento).Value.Length > 0)
                {
                    //Fecha de nacimiento:
                    _cElemento = "Fecha de nacimiento";
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:DateOfBirth")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_de_nacimiento).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais_del_telefono).Value.Length > 0)
                {
                    //Código País del teléfono:
                    _cElemento = "Código país";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:PhoneContactInputSet:CountryCelularTelephone", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais_del_telefono).Value);
                }

                //falta teléfono celular
                _cElemento = "Teléfono celular";
                _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:PhoneContactInputSet:CellPHone")).SendKeys("978855614");

                //falta correo principal
                _cElemento = "Correo electrónico principal";
                _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:ContactEmailsInputSet:PrimaryEmailTypeExt", "Personal");

                //falta correo personal
                _cElemento = "Correo electrónico personal";
                _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:PolicyContactRoleNameInputSet:ContactEmailsInputSet:EmailAddress1")).SendKeys("mi nueva dirección");

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_direccion).Value.Length > 0)
                {
                    //Tipo de dirección:
                    _cElemento = "Tipo de dirección";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressType", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_direccion).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais).Value.Length > 0)
                {
                    //País:
                    _cElemento = "País";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_Country", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.departamento).Value.Length > 0)
                {
                    //Departamento:
                    _cElemento = "Departamento";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_Department", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.departamento).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.provincia).Value.Length > 0)
                {
                    //Provincia:
                    _cElemento = "Provincia";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_Province", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.provincia).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.distrito).Value.Length > 0)
                {
                    //Distrito:
                    _cElemento = "Distrito";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_District", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.distrito).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_calle).Value.Length > 0)
                {
                    //Tipo de calle:
                    _cElemento = "Tipo de calle";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_StreetType", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_calle).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre_de_la_calle).Value.Length > 0)
                {
                    //Nombre de la calle:
                    _cElemento = "Nombre de la calle";
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressLine1")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre_de_la_calle).Value);
                }

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.numero).Value.Length > 0)
                {
                    //Número:
                    _cElemento = "Número";
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressLine2")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.numero).Value);
                }

                //Finalizar formulario:
                //Clic en botón Aceptar:
                _cElemento = "Botón Aceptar";
                _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:Update")).Click();
                _Funciones.Esperar(3);

                //Clic en Cotización:
                _cElemento = "Botón Cotización";
                _driverGlobal.FindElement(By.Id("PolicyChangeWizard:LOBWizardStepGroup:PolicyChangeWizard_PolicyInfoScreen:JobWizardToolbarButtonSet:QuoteOrReview")).Click();
                _Funciones.Esperar(10);

                //Clic en botón Detalles:
                _cElemento = "Botón Detalles";
                _driverGlobal.FindElement(By.Id("UWBlockProgressIssuesPopup:IssuesScreen:DetailsButton")).Click();
                _Funciones.Esperar(2);
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error en formulario Editar Cuenta: " + Ex.Message + " " + _cElemento, Ex); }
        }

        //Pasos en Ventana Análisis de Riesgo:
        private Boolean AnalisisDeRiesgos()
        {
            try
            {
                //Ventana de aprobación de bloqueantes:
                if (_Funciones.ExisteElemento(_driverGlobal, "PolicyChangeWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:1:UWIssueRowSet:ShortDescription", 2))
                {
                    //Marcar check de aprobación:
                    _cElemento = "Check bloqueante de cotización";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:1:UWIssueRowSet:_Checkbox")).Click();

                    if (_Funciones.ExisteElemento(_driverGlobal, "PolicyChangeWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:Approve"))
                    {
                        //Clic en botón Aprobar:
                        _cElemento = "Botón Aprobar";
                        _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:Approve")).Click();
                        _Funciones.Esperar(2);

                        //Clic en opción SI:
                        _cElemento = "Permitir edición";
                        _driverGlobal.FindElement(By.Id("RiskApprovalDetailsPopup:0:IssueDetailsDV:UWApprovalLV:EditBeforeBind_true")).Click();

                        //Clic en Botón Aceptar:
                        _cElemento = "Botón Aceptar";
                        _driverGlobal.FindElement(By.Id("RiskApprovalDetailsPopup:Update")).Click();
                        _Funciones.Esperar(2);
                    }
                    else { return false; }
                }
                else { return false; }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error en análisis de riesgos: " + _cElemento, Ex); }
            return true;
        }

        private void ConfirmarTrabajo()
        {
            try
            {
                //Clic en Cotización:
                _cElemento = "Botón Cotización";
                _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:JobWizardToolbarButtonSet:QuoteOrReview")).Click();
                _Funciones.Esperar(10);

                //Clic en Confirmar trabajo:
                _cElemento = "Botón Confirmar trabajo";
                _driverGlobal.FindElement(By.Id("PolicyChangeWizard:PolicyChangeWizard_QuoteScreen:JobWizardToolbarButtonSet:BindPolicyChange")).Click();
                _Funciones.VentanaMensajeWeb(_driverGlobal);
                _Funciones.Esperar(5);

                //Cambio de póliza Completada:
                _cElemento = "Cambio completado";
                _driverGlobal.FindElement(By.Id("JobComplete:JobCompleteScreen:JobCompleteDV:ViewPolicy")).Click();
                _Funciones.Esperar(4);
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al confirmar trabajo: " + Ex.Message + " " + _cElemento, Ex); }
        }
        #endregion
    }
}
