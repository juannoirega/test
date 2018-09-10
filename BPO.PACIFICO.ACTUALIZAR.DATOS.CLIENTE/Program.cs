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
                catch (Exception Ex)
                {
                    CambiarEstadoTicket(oTicket, _oMesaControl, Ex.Message);
                    LogFailStep(30, Ex);
                    _Funciones.CerrarDriver(_driverGlobal);
                    return;
                    //_oRobot.SaveTicketNextState(_Funciones.MesaDeControl(oTicket,Ex.Message), _oRobot.GetNextStateAction(oTicket).First(o => o.DestinationStateId == Convert.ToInt32(_EstadoSiguiente)).Id);
                }
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
                int[] oCampos = new int[] { eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado,
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

                    if(_Funciones.ExisteElemento(_driverGlobal, "PolicyChangeWizard:OfferingScreen:OfferingSelection",2))
                        //_driverGlobal

                    //Clic en Nombre del asegurado:
                    _cElemento = "Nombre del Asegurado";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:LOBWizardStepGroup:PolicyChangeWizard_PolicyInfoScreen:PolicyChangeWizard_PolicyInfoDV:AccountInfoInputSet:Name")).Click();
                    _Funciones.Esperar(7);

                    //Método para actualizar datos:


                    //Botón Cotización:
                    _cElemento = "Botón Cotización";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:OfferingScreen:JobWizardToolbarButtonSet:QuoteOrReview")).Click();
                    _Funciones.Esperar(5);
                }
                else if (_cLinea == _cLineaRRGG)
                {
                    //Seleccionar tipo de endoso:
                    _cElemento = "Tipo de endoso";
                    _Funciones.SeleccionarCombo(_driverGlobal, "StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:TypeReason", _Funciones.ObtenerValorDominio(oTicketDatos, Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_proceso).Value)));
                    _Funciones.Esperar();

                    //Seleccionar subtipo de endoso:
                    _cElemento = "Subtipo de endoso";
                    _Funciones.SeleccionarCombo(_driverGlobal, "StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:Description", _Funciones.ObtenerValorDominio(oTicketDatos, Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_anu_motivo).Value)));
                    _Funciones.Esperar();

                    //Ingresar comentarios adicionales:
                    _cElemento = "Comentarios adicionales";
                    _driverGlobal.FindElement(By.Id("StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:Comments")).SendKeys(_cComentariosAdicionales);

                    //Clic en Siguiente:
                    _cElemento = "Botón Siguiente";
                    _driverGlobal.FindElement(By.Id("StartPolicyChange:StartPolicyChangeScreen:NewPolicyChange")).Click();
                    _Funciones.Esperar(3);

                    //Clic en Siguiente:
                    _cElemento = "Botón Siguiente:";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Next")).Click();
                    _Funciones.Esperar();

                    //Clic en Nombre:
                    _cElemento = "Nombre del Asegurado";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:LOBWizardStepGroup:PolicyChangeWizard_PolicyInfoScreen:PolicyChangeWizard_PolicyInfoDV:AccountInfoInputSet:Name")).Click();
                    _Funciones.Esperar(5);

                    //Método para actualizar datos:


                    //Clic en botón Cotización:
                    _cElemento = "Botón Cotización";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:LOBWizardStepGroup:PolicyChangeWizard_PolicyInfoScreen:JobWizardToolbarButtonSet:QuoteOrReview")).Click();
                    _Funciones.Esperar(20);
                }
                else if (_cLinea == _cLineaAlianzas)
                {

                }
                else if (_cLinea == _cLineaLLPP)
                {
                    //Buscar Documento:
                    _cElemento = "Buscar por documento";
                    _Funciones.BuscarDocumentoPolicyCenter(_driverGlobal, 
                        oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.dni).Value.Length == 0? 
                        oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.ruc).Value: 
                        oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.dni).Value);

                    //Clic en Nro. de Cuenta:
                    _cElemento = "Clic en Nro. de Cuenta";
                    _driverGlobal.FindElement(By.Id("ContactFile_AccountsSearch:AssociatedAccountsLV:0:AccountNumber")).Click();
                    _Funciones.Esperar(27);

                    //Clic en Editar Cuenta:
                    _cElemento = "Editar cuenta";
                    _driverGlobal.FindElement(By.Id("AccountFile_Summary:AccountFile_SummaryScreen:EditAccount")).Click();
                    _Funciones.Esperar(4);

                    if (FormularioEditarCuenta(oTicketDatos))
                    {
                        //Clic en Actualizar:
                        _cElemento = "Clic en Actualizar";
                        _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:Update")).Click();
                        _Funciones.Esperar(22);
                    }  
                }


                //Modificar dirección:
                if (Convert.ToString(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_calle).Value).Length > 0)
                {
                    _cElemento = "Tipo de calle:";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_StreetType",
                        _Funciones.ObtenerValorDominio(oTicketDatos, Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_calle).Value)));

                    _cElemento = "Nombre de la calle:";
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressLine1")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre_de_la_calle).Value);

                    _cElemento = "Número";
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressLine2")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.numero_de_dni).Value); //poner numero de la calle.

                    _cElemento = "Referencia";
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:PolicyContactRoleDetailsCV:PolicyContactDetailsDV:AddressExtInputSet:Address_AddressLine2")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre_de_la_calle).Value); //poner Referencia

                    //Clic en botón Aceptar:
                    _cElemento = "Botón Aceptar";
                    _driverGlobal.FindElement(By.Id("EditPolicyContactRolePopup:ContactDetailScreen:Update")).Click();
                    _Funciones.Esperar(3);
                }


                //Clic en botón Detalles:
                _cElemento = "Botón Detalles";
                _driverGlobal.FindElement(By.Id("UWBlockProgressIssuesPopup:IssuesScreen:DetailsButton")).Click();
                _Funciones.Esperar(2);

                if (_Funciones.ExisteElemento(_driverGlobal, "PolicyChangeWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:1:UWIssueRowSet:ShortDescription", 2))
                {
                    //Marcar check de aprobación:
                    _cElemento = "Check bloqueante de cotización";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:1:UWIssueRowSet:_Checkbox")).Click();

                    //Clic en botón Aprobar:
                    _cElemento = "Botón Aprobar";
                    _driverGlobal.FindElement(By.Id("PolicyChangeWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:Approve")).Click();
                }




                //Clic en Cambiar a:
                _cElemento = "Cambiar a:";
                _driverGlobal.FindElement(By.Id("PolicyChangeWizard:LOBWizardStepGroup:PolicyChangeWizard_PolicyInfoScreen:PolicyChangeWizard_PolicyInfoDV:AccountInfoInputSet:PolicyAddressDisplayAutoPersonalInputSet:ChangePolicyDeliveryAddressButton:ChangePolicyDeliveryAddressButtonMenuIcon")).Click();

                //Clic en Editar dirección actual:
                _cElemento = "Editar dirección actual";
                _driverGlobal.FindElement(By.Id("PolicyChangeWizard:LOBWizardStepGroup:PolicyChangeWizard_PolicyInfoScreen:PolicyChangeWizard_PolicyInfoDV:AccountInfoInputSet:PolicyAddressDisplayAutoPersonalInputSet:ChangePolicyDeliveryAddressButton:EditDeliveryAddressMenuItem")).Click();

                //Enviar Ticket al siguiente estado:
                CambiarEstadoTicket(oTicketDatos, _oTicketHijo);
            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
                throw new Exception(Ex.Message + " :" + _cElemento, Ex);
            }
        }

        //Inicia acción Cambiar Póliza:
        private void IniciarCambioPoliza(Ticket oTicketDatos)
        {
            try
            {
                //Ingresar mediante Menú Acciones:
                _cElemento = "Menú Acciones";
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions")).Click();

                //Hacer clic en Cambiar Póliza:
                _cElemento = "Opción Cambiar Póliza";
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions:PolicyFileMenuActions_NewWorkOrder:PolicyFileMenuActions_ChangePolicy")).Click();
                _Funciones.Esperar(3);

                //Formulario registro de endoso:
                _cElemento = "Fecha efectiva del cambio";
                //Validar fecha:
                if (ValidarFechaSolicitud(oTicketDatos, Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_efectiva).Value)))
                    _driverGlobal.FindElement(By.Id("StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:EffectiveDate")).
                        SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_efectiva).Value);
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al iniciar el cambio de la póliza: " + _cElemento, Ex); }
        }

        //Formulario inicial para el Cambio de Póliza:
        private void FormularioCambioPoliza(Ticket oTicketDatos)
        {
            try
            {
                //Seleccionar tipo de complejidad:
                _cElemento = "Tipo de Complejidad";
                _Funciones.SeleccionarCombo(_driverGlobal, "StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:TypeReason", _Funciones.ObtenerValorDominio(oTicketDatos, Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_producto).Value))); //Campo tipo de complejidad
                _Funciones.Esperar();

                //Seleccionar motivo del endoso:
                _cElemento = "Motivo del endoso";
                _Funciones.SeleccionarCombo(_driverGlobal, "StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:Description", _Funciones.ObtenerValorDominio(oTicketDatos, Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_anu_motivo).Value)));
                _Funciones.Esperar();

                //Ingresar comentarios adicionales:
                _cElemento = "Comentarios adicionales";
                _driverGlobal.FindElement(By.Id("StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:Comments")).SendKeys(_cComentariosAdicionales);

                //Clic en Siguiente:
                _cElemento = "Botón Siguiente";
                _driverGlobal.FindElement(By.Id("StartPolicyChange:StartPolicyChangeScreen:NewPolicyChange")).Click();
                _Funciones.Esperar(6);
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al llenar formulario Cambio de Póliza: " + _cElemento, Ex); }
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

        private Boolean FormularioEditarCuenta(Ticket oTicketDatos)
        {
            bool bEditado = true;
            try
            {

                if (oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre).Value.Length > 0)
                {
                    //Nombre:
                    _cElemento = "Nombre";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:FirstName")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.nombre).Value);
                }
                    
                if(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.apellido_paterno).Value.Length > 0)
                {
                    //Apellido Paterno:
                    _cElemento = "Apellidos paterno";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:LastName")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.apellido_paterno).Value);
                }

                if(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.apellido_materno).Value.Length > 0)
                {
                    //Apellido Materno:
                    _cElemento = "Apellidos paterno";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:SecondLastNameExt")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.apellido_materno).Value);
                }
                    
                if(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais_de_procedencia).Value.Length > 0)
                {
                    //País de procedencia:
                    _cElemento = "País de procedencia";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:ForeignCountry", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais_de_procedencia).Value);
                }

                if(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_de_nacimiento).Value.Length > 0)
                {
                    //Fecha de nacimiento:
                    _cElemento = "Fecha de nacimiento";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:ContactNameInputSet:ContactBasicInformationInputSet:DateOfBirth")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_de_nacimiento).Value);
                }

                if(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_direccion).Value.Length > 0)
                {
                    //Tipo de dirección:
                    _cElemento = "Tipo de dirección";
                    _Funciones.SeleccionarCombo(_driverGlobal, "EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_AddressType", oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_direccion).Value);
                }

                if(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.pais).Value.Length > 0)
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
                    _cElemento = "Provincia";
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
                    //Nombre de la calle:
                    _cElemento = "Nombre de la calle";
                    _driverGlobal.FindElement(By.Id("EditAccountPopup:EditAccountScreen:AddressExtInputSet:Address_AddressLine2")).SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.numero).Value);
                }
            }
            catch (Exception Ex) { bEditado = false; throw new Exception("Ocurrió un error al llenar formulario Editar Cuenta: " + _cElemento, Ex) ; }

            return bEditado;
        }
    }
}
