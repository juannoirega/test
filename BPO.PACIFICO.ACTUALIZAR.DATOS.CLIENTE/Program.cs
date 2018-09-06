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
        private static string[] Usuarios; 
        private static int _nIndice;
        private static string _cElemento = string.Empty;
        private static string _cLineaAutos = string.Empty;
        private static string _cLineaRRGG = string.Empty;
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
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al obtener parámetros del Robot: "+ Ex.Message, Ex); }
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
            if (!ValidarVacios(oTicket))
            {
                _nIndice = 1;
                for (int i = 0; i < _nIntentosPolicyCenter; i++)
                {
                    _Funciones.AbrirSelenium(ref _driverGlobal);
                    _Funciones.NavegarUrlPolicyCenter(_driverGlobal, _cUrlPolicyCenter);
                    Credenciales();
                    _Funciones.LoginPolicyCenter(_driverGlobal, Usuarios[0], Usuarios[1]);
                    if (_Funciones.ExisteElemento(_driverGlobal, "TabBar:PolicyTab_arrow", _nIntentosPolicyCenter))
                    {
                        break;
                    }
                    _Funciones.CerrarDriver(_driverGlobal);
                    _nIndice += 1;
                }
                ActualizarPolicyCenter(oTicket);
            }
        }

        //Obtiene los usuarios con sus respectivas contraseñas:
        private string[] Credenciales()
        {
            try
            {
                //Usuario y contraseña de Autos:
                Usuarios = _oRobot.GetValueParamRobot("AccesoPCyCM_" + _nIndice).ValueParam.Split(',');
                return Usuarios;
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
                //Busca datos de Póliza por Nro. de Póliza:
                _cElemento = "Buscar póliza";
                _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_nro).Value);

                //Ingresar mediante Menú Acciones:
                _cElemento = "Menú Acciones";
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions")).Click();

                //Hacer clic en Cambiar Póliza:
                _cElemento = "Opción Cambiar Póliza";
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions:PolicyFileMenuActions_NewWorkOrder:PolicyFileMenuActions_ChangePolicy")).Click();
                _Funciones.Esperar(3);

                //Formulario registro de endoso:
                _cElemento = "Fecha efectiva del cambio";
                _driverGlobal.FindElement(By.Id("StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:EffectiveDate")).
                    SendKeys(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.fecha_hora_de_email).Value);

                if (_cLinea == _cLineaAutos)
                {
                    _cElemento = "Tipo de Complejidad";
                    _Funciones.SeleccionarCombo(_driverGlobal, "StartPolicyChange:StartPolicyChangeScreen:StartPolicyChangeDV:TypeReason", _Funciones.ObtenerValorDominio(oTicketDatos, Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_producto).Value)));
                }
                else if (_cLinea == _cLineaRRGG)
                {

                }
            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
                throw new Exception(Ex.Message + " :" +_cElemento, Ex);
            }
        }
    }
}
