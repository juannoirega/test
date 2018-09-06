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
        private static BaseRobot<Program> _robot = null;
        private static IWebDriver _driverGlobal = new InternetExplorerDriver();
        string campo = string.Empty;
        string texto = string.Empty;
        string Acceso = string.Empty;
        private static Functions _Funciones;

        static string[] _valoresTickets = new string[10];
        static string[] _valoresTickets_Ident = new string[10];


        List<TicketValue> ticketValue = null;

        #region ParametrosRobot
        private string _urlContactManager = string.Empty;
        private string _usuarioContactManager = string.Empty;
        private string _contraseñaContactManager = string.Empty;
        private string _EstadoError = string.Empty;
        private string _EstadoSiguiente = string.Empty;

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

            foreach (Ticket ticket in _robot.Tickets)
            {
                try
                {
                    ProcesarTicket(ticket);
                }
                catch (Exception ex)
                {
                    LogFailStep(30, ex);

                    _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, ex.Message), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == Convert.ToInt32(_EstadoSiguiente)).Id);
                }

            }


        }

        private void ProcesarTicket(Ticket ticket)
        {

            GetParameterRobots();
            AbrirSelenium();
            NavegarUrl();
            Login();
            AccedientoContactManager(ticket);



        }

        private void AccedientoContactManager(Ticket ticket)
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
                _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, "Registro no Encontrado"), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == Convert.ToInt32(_EstadoSiguiente)).Id);
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

                ticket = _robot.Tickets.FirstOrDefault();

                ticketValue = _robot.GetDataQueryTicketValue().Where(a => a.TicketId == ticket.Id).ToList();


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

                _robot.SaveTicketNextState(ticket, Convert.ToInt32(_EstadoSiguiente));

            }
            catch (Exception ex)
            {

                throw;
            }

        }


        private void GetParameterRobots()
        {
            try
            {
                _urlContactManager = _robot.GetValueParamRobot("URLContactManager").ValueParam;
                _usuarioContactManager = _robot.GetValueParamRobot("UsuarioContactManager").ValueParam;
                _contraseñaContactManager = _robot.GetValueParamRobot("PasswordContactManager").ValueParam;
                _EstadoError = _robot.GetValueParamRobot("EstadoError").ValueParam;
                _EstadoSiguiente = _robot.GetValueParamRobot("EstadoSiguiente").ValueParam;

            }
            catch (Exception ex)
            {
                throw new Exception("Error al Obtener los parametros del robot", ex);
            }
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

            try
            {
                //LogInfoStep(5);//id referencial msje Log "Iniciando acceso al sitio policenter"
                _Funciones.NavegarUrlPolicyCenter(_driverGlobal, _urlContactManager);
            }
            catch (Exception ex)
            {
                throw new Exception("No se puede acceder al sitio policycenter", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Finalizando acceso al sitio policenter"


        }

        private void Login()
        {

            try
            {
                //LogInfoStep(5);//id referencial msje Log "Iniciando login policenter"

                _Funciones.LoginPolicyCenter(_driverGlobal, _usuarioContactManager, _contraseñaContactManager);
            }
            catch (Exception ex)
            {
                throw new Exception("No se puede acceder al sistema policycenter", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Finalizacion login policenter"


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



    }
}
