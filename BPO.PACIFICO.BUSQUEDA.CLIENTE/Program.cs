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
                        _Funciones.Esperar(3);
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
                        _Funciones.Esperar(3);
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
                        _Funciones.Esperar(3);
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
                        _Funciones.Esperar(3);
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
                CapturarDatos(ticket);
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

        private void CapturarDatos(Ticket ticket)
        {
            try
            {

                //País de procedencia
                InsertarValoresFielt(ticket, 1067, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:CountryOfOriginExt"));


                if (_valoresTickets[0] != "")
                {
                    //Nombre (s)
                    InsertarValoresFielt(ticket, 1084, ExtraerDatosXPath(0, 4, "textBox"));
                    //Apellido Paterno
                    InsertarValoresFielt(ticket, 1070, ExtraerDatosXPath(0, 5, "textBox"));
                    //Apellido Materno
                    InsertarValoresFielt(ticket, 1069, ExtraerDatosXPath(0, 6, "textBox"));
                    //Apellido Materno
                    InsertarValoresFielt(ticket, 1072, ExtraerDatosXPath(0, 8, "textBox"));
                    //Prefijo
                    InsertarValoresFielt(ticket, 1087, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:Prefix"));
                    //Fecha de nacimiento
                    InsertarValoresFielt(ticket, 1083, ExtraerDatosXPath(5, 5, "textBox"));
                    //Sexo
                    InsertarValoresFielt(ticket, 1088, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:ABPersonVendorInputSet:Gender"));
                    //Estado civil                                         
                    InsertarValoresFielt(ticket, 1089, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:ABPersonVendorInputSet:MaritalStatus"));
                    //Correo Personal;                                     
                    InsertarValoresFielt(ticket, 1084, ExtraerDatosXPath(0, 41, "textBox"));
                    //Distrito
                    InsertarValoresFielt(ticket, 1078, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_District"));

                }

                if (_valoresTickets[1] != "")
                {
                    //Razón Social
                    InsertarValoresFielt(ticket, 1085, ExtraerDatosXPath(0, 4, "textBox"));
                    //Nombre comercial 
                    InsertarValoresFielt(ticket, 1086, ExtraerDatosXPath(0, 5, "textBox"));
                    //Fecha de inicio de Actividades;
                    InsertarValoresFielt(ticket, 1094, ExtraerDatosXPath(6, 5, "textBox"));
                    //Actividad económica
                    InsertarValoresFielt(ticket, 1095, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:EconomicSectorActivityInputSet:EconomicSubSectorExt"));
                    //Sector Económico                                     
                    InsertarValoresFielt(ticket, 1096, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:EconomicSectorActivityInputSet:EconomicSectorExt"));
                    //Correo Personal;
                    InsertarValoresFielt(ticket, 1084, ExtraerDatosXPath(0, 38, "textBox"));
                    //Distrito
                    InsertarValoresFielt(ticket, 1078, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:abc"));

                }


                //Teléfono principal
                //InsertarValoresFielt(ticket, 1067, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:ABPhoneDetailsInputSet:PrimaryPhone"));

                //País del teléfono
                InsertarValoresFielt(ticket, 1090, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:ABPhoneDetailsInputSet:CellPhoneCountry"));
                //Indicativo(código de área)                           
                InsertarValoresFielt(ticket, 1092, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:ABPhoneDetailsInputSet:HomeAreaCodeExtPeru"));
                //Teléfono de Casa  ContactDetail:ABContactDetailScreen:ContactBasicsDV:ABPhoneDetailsInputSet:OtherAreaCodeExtPeru
                InsertarValoresFielt(ticket, 1093, ExtraerDatosXPath(0, 18, "textBox"));

                //País
                InsertarValoresFielt(ticket, 1075, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_Country"));
                //Departamento                                         ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_Country
                InsertarValoresFielt(ticket, 1076, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_Department"));
                //Provincia
                InsertarValoresFielt(ticket, 1077, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_Province"));
                //Tipo de calle                                         ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_District
                InsertarValoresFielt(ticket, 1079, OptenerTextoSelect("ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_StreetType"));
                //Nombre de la calle                                   ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_StreetType
                //                                                     ContactDetail:ABContactDetailScreen:ContactBasicsDV:PrimaryAddressInputSet:AddressOwnerInputSet:Address_StreetType     
                InsertarValoresFielt(ticket, 1080, ExtraerDatosXPath(0, 53, "textBox"));
                //Número
                InsertarValoresFielt(ticket, 1081, ExtraerDatosXPath(0, 54, "textBox"));
                //Referencia
                InsertarValoresFielt(ticket, 1082, ExtraerDatosXPath(0, 55, "textBox"));

                _robot.SaveTicketNextState(ticket, Convert.ToInt32(_EstadoSiguiente));
            }

            catch (Exception ex)
            {

                _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, ex.Message), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == Convert.ToInt32(_EstadoSiguiente)).Id);
            }

        }

        public void InsertarValoresFielt(Ticket ticket, Int32 idFields, String valor)
        {

            ticket.TicketValues.Add(new TicketValue
            {
                FieldId = idFields,
                Value = valor,
                IsEncrypted = false,
                ClonedValueOrder = null
            });
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

        
        public String ExtraerDatosXPath(int index, int posicion, string clase)
        {
            try
            {
                //IWebElement _element = null; ;
                string xPath = "//*[@id='ContactDetail:ABContactDetailScreen:ContactBasicsDV:" + index + "']/tbody/tr[" + posicion + "]/td[5]/input[@class='" + clase + "']";
                return _driverGlobal.FindElement(By.XPath(xPath)).GetAttribute("value");
            }
            catch (Exception)
            {
                return "";
            }

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

        public string OptenerTextoSelect(String name)
        {
            SelectElement elemen = new SelectElement(_driverGlobal.FindElement(By.Name(name)));
            return elemen.SelectedOption.Text;


        }




    }
}

