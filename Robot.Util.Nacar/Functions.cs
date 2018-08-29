using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using BPO.Framework.BPOCitrix;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.IE;
using System.Threading;
using Everis.Ees.Entities;
using everis.Ees.Proxy.Services;
using OpenQA.Selenium.Interactions;

namespace Robot.Util.Nacar
{
    public class Functions
    {
        #region "PARÁMETROS"
        private static string _cRutaGeckodriver = string.Empty;
        private static string _cRutaFirefox = string.Empty;
        private static string _cBPMWebDriver = string.Empty;
        private static string _cGeckodriver = string.Empty;
        #endregion

        //Registro en BPM:
        public void IngresarBPM(IWebDriver oDriver, string Url, string Usuario, string Contraseña)
        {
            try
            {
                oDriver.Navigate().GoToUrl(Url);
                Esperar();
                VentanaWindows(oDriver, Usuario, Contraseña);
                oDriver.Manage().Window.Maximize();
            }
            catch (Exception Ex)
            {
                throw new Exception("Error de acceso al sistema OnBase.", Ex);
            }
        }

        public void AbrirSelenium(ref IWebDriver _driver)
        {
            _driver = new InternetExplorerDriver();
            _driver.Manage().Window.Maximize();
        }

        public void NavegarUrlPolicyCenter(IWebDriver _driver, string url)
        {
            _driver.Url = url;
            _driver.Navigate().GoToUrl("javascript:document.getElementById('overridelink').click()");
            _driver.Manage().Window.Maximize();
            Esperar(1);
        }

        public Ticket MesaDeControl(Ticket ticket, string motivo)
        {
            if (ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.trazabilidad) == null)
            {
                ticket.TicketValues.Add(new TicketValue
                {
                    FieldId = eesFields.Default.trazabilidad,
                    TicketId = ticket.Id,
                    Value = motivo,
                    CreationDate = DateTime.Now,
                    ClonedValueOrder = null
                });
            }
            else
                ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.trazabilidad).Value = motivo;
            return ticket;
        }

        public List<TicketValue> ValuesPadre(Ticket ticket, Default.Container container)
        {
            List<TicketValue> values = new List<TicketValue>();

            if (ticket.ParentId != 0 && ticket.ParentId != null)
                values.AddRange(ValuesPadre(container.Tickets.Expand(tv => tv.TicketValues).Where(o => o.Id == ticket.ParentId).First(), container));

            values.AddRange(ticket.TicketValues);

            return values;
        }

        public void NavegarUrlPortalBcp(IWebDriver _driver, string url)
        {
            _driver.Url = url;
            _driver.Manage().Window.Maximize();
            Esperar(1);
        }

        public void LoginPolicyCenter(IWebDriver _driver, string usuario, string contraseña)
        {
            _driver.SwitchTo().Frame(_driver.FindElement(By.Id("top_frame")));

            _driver.FindElement(By.Id("Login:LoginScreen:LoginDV:username")).SendKeys(usuario);
            _driver.FindElement(By.Id("Login:LoginScreen:LoginDV:password")).SendKeys(contraseña);
            _driver.FindElement(By.Id("Login:LoginScreen:LoginDV:submit")).SendKeys(Keys.Enter);
            Esperar(1);
        }
        public void LoginPortalBcp(IWebDriver _driver, string usuario, string contraseña)
        {
            _driver.FindElement(By.Id("ctl00_MainContent_txtUsuario")).SendKeys(usuario);
            _driver.FindElement(By.Id("ctl00_MainContent_txtPassword")).SendKeys(contraseña);
            //Para pruebas colocar punto interrupcion para ingresar captcha manualmente
            //FALTA IMPLEMENTAR EL CAPTCHA
        }

        public void BuscarPolizaPolicyCenter(IWebDriver _driver, string numeroPoliza)
        {
            _driver.FindElement(By.Id("TabBar:PolicyTab_arrow")).Click();
            _driver.FindElement(By.Id("TabBar:PolicyTab:PolicyTab_PolicyRetrievalItem")).SendKeys(numeroPoliza);
            _driver.FindElement(By.Id("TabBar:PolicyTab:PolicyTab_PolicyRetrievalItem")).SendKeys(Keys.Enter);
            Esperar(5);
        }

        public void BuscarPolizaPortalBcp(IWebDriver _driver, string numeroPoliza)
        {
            IWebElement element;
            Actions action = new Actions(_driver);
            element = _driver.FindElement(By.XPath("//span[contains(.,'Modificaciones')]"));
            Esperar(2);
            action.MoveToElement(element).MoveToElement(_driver.FindElement(By.XPath("//a[contains(.,'Registrar Modificaciones')]"))).Click().Build().Perform();
            Esperar(2);
            _driver.SwitchTo().DefaultContent();
            _driver.SwitchTo().Frame(_driver.FindElement(By.Id("ifrmApp")));
            _driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_txtNroPolizaMod")).SendKeys(numeroPoliza);
            _driver.FindElement(By.XPath("//input[contains(@id,'ctl00_ContentPlaceHolder1_btnConsultar')]")).Click();
            Esperar(2);
        }

        public string ObtenerValorElemento(IWebDriver _driver, string idElemento, string type = "id")
        {
            switch (type.ToLower())
            {
                case "id":
                    return _driver.FindElement(By.Id(idElemento)).Text;
                case "xpath":
                    return _driver.FindElement(By.XPath(idElemento)).Text;
                case "linktext":
                    return _driver.FindElement(By.LinkText(idElemento)).Text;
                case "name":
                    return _driver.FindElement(By.Name(idElemento)).Text;
                case "tagname":
                    return _driver.FindElement(By.TagName(idElemento)).Text;
                case "classname":
                    return _driver.FindElement(By.ClassName(idElemento)).Text;
                case "partiallinktext":
                    return _driver.FindElement(By.PartialLinkText(idElemento)).Text;
                default:
                    return null;
            }
        }

        public void SeleccionarCombo(IWebDriver _driver, string idElement, string valorComparar)
        {
            try
            {
                //Anina: Obtiene nombre del Driver
                IList<IWebElement> oOption;
                Type oTypeDriver = _driver.GetType();
                //Obtiene lista de opciones según Webdriver:
                if (oTypeDriver.Name == _cBPMWebDriver)
                {
                    oOption = _driver.FindElements(By.XPath(idElement));
                }
                else
                {
                    oOption = _driver.FindElement(By.Id(idElement)).FindElements(By.XPath("id('" + idElement + "')/option"));
                }

                for (int i = 0; i < oOption.Count; i++)
                {
                    if (oOption[i].Text.ToUpperInvariant().Equals(valorComparar))
                    {
                        oOption[i].Click();
                        Esperar();
                        break;
                    }
                }
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un error al seleccionar una opción.", Ex);
            }

        }

        public string ObtenerValorDominio(Ticket ticket, int idCampoDominio)
        {
            var container = ODataContextWrapper.GetContainer();
            try
            {
                if (ticket != null)
                    return container.DomainValues.FirstOrDefault(p => p.Id == idCampoDominio).Value.Trim().ToUpperInvariant();
            }
            catch
            {
                return null;
            }
            return null;
        }

        //Método para hacer pausa en segundos:
        public void Esperar(double nTiempo = 1)
        {
            Thread.Sleep(1000 * Convert.ToInt32(nTiempo));
        }

        //Anina: Ingresar usuario y contraseña en ventana windows.
        public void VentanaWindows(IWebDriver oDriver, string cUsuario, string cContraseña)
        {
            var oAlert = oDriver.SwitchTo().Alert();
            oAlert.SendKeys(Keys.Clear);
            oAlert.SendKeys(cUsuario + Keys.Tab + cContraseña);
            Esperar();
            oAlert.Accept();
            Esperar(18);
        }

        private static FirefoxDriver GetFirefoxDriver()
        {
            var oDriver = new FirefoxDriver(GetFirefoxDriverService(), GetFirefoxOptions());
            return oDriver;
        }

        private static FirefoxDriverService GetFirefoxDriverService()
        {
            try
            {
                var service = FirefoxDriverService.CreateDefaultService(_cRutaGeckodriver, _cGeckodriver);
                service.FirefoxBinaryPath = _cRutaFirefox;
                return service;
            }
            catch (Exception Ex)
            {
                throw new Exception("No se encontró la ruta para el archivo geckodriver.exe", Ex);
            }
        }

        private static FirefoxOptions GetFirefoxOptions()
        {
            var oOptions = new FirefoxOptions();

            oOptions.SetPreference("browser.download.manager.focusWhenStarting", false);
            oOptions.SetPreference("browser.helperApps.alwaysAsk.force", false);
            oOptions.SetPreference("browser.download.manager.alertOnEXEOpen", false);
            oOptions.SetPreference("browser.download.manager.closeWhenDone", false);
            oOptions.SetPreference("browser.download.manager.showAlertOnComplete", false);
            oOptions.SetPreference("browser.download.manager.useWindow", false);
            oOptions.SetPreference("services.sync.prefs.sync.browser.download.manager.showWhenStarting", false);
            oOptions.SetPreference("browser.cache.disk.enable", false);
            oOptions.SetPreference("browser.cache.memory.enable", false);
            oOptions.SetPreference("browser.cache.offline.enable", false);
            oOptions.SetPreference("network.http.use-cache", false);
            oOptions.SetPreference("network.auth.subresource-http-auth-allow", 2);
            oOptions.AcceptInsecureCertificates = true;
            oOptions.AddArgument("--trustAllSSLCertificates");

            oOptions.Profile = GetFirefoxProfile();
            return oOptions;
        }

        private static FirefoxProfile GetFirefoxProfile()
        {
            var oProfile = new FirefoxProfile();
            oProfile.AcceptUntrustedCertificates = true;
            oProfile.AssumeUntrustedCertificateIssuer = true;
            return oProfile;
        }

        public void CerrarDriver(IWebDriver oDriver)
        {
            if (oDriver != null)
            {
                oDriver.Close();
                oDriver.Quit();
            }
        }

        //Anina: Crea una nueva instancia para Firefox.
        public void InstanciarFirefoxDriver(ref IWebDriver oDriver, string RutaGeckodriver, string RutaFirefox, string BPMWebDriver, string Geckodriver)
        {
            try
            {
                _cRutaGeckodriver = RutaGeckodriver;
                _cRutaFirefox = RutaFirefox;
                _cBPMWebDriver = BPMWebDriver;
                _cGeckodriver = Geckodriver;
                oDriver = GetFirefoxDriver();
                Esperar();
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un error al abrir navegador Firefox.", Ex);
            }
        }

        //Anina: Selecciona elemento de una lista.
        public void SeleccionarListBox(IWebDriver oDriver, string cElemento, string cValorComparar)
        {
            try
            {
                Actions oAcciones = new Actions(oDriver);

                //Obtiene lista de opciones:
                IList<IWebElement> oOption = oDriver.FindElements(By.XPath(cElemento));

                for (int i = 0; i < oOption.Count; i++)
                {
                    if (oOption[i].Text.ToUpperInvariant().Equals(cValorComparar))
                    {
                        oAcciones.MoveToElement(oOption[i]).Click();
                        oAcciones.Build().Perform();
                        Esperar();
                        break;
                    }
                }
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un error al seleccionar una opción de la lista.", Ex);
            }
        }

        //Valida campos vacíos en TicketValues:
        public Boolean ValidarCamposVacios(Ticket oTicket, int [] oCampos)
        {
            foreach (int nCampo in oCampos)
                if (String.IsNullOrWhiteSpace(oTicket.TicketValues.FirstOrDefault(o => o.FieldId == nCampo).Value.Trim()))
                    return false;
            return true;
        }

        //Anina: Verifica si existe elemento web.
        public Boolean ExisteElemento(IWebDriver oDriver, string cIdElemento, int nIntentos)
        {
            for (int i = 0; i < nIntentos; i++)
            {
                try
                {
                    oDriver.FindElement(By.Id(cIdElemento));
                    return true;
                }
                catch (NoSuchElementException)
                {
                    Esperar();
                }
            }
            return false;
        }

        //Verifica que el formulario BPM se haya registrado correctamente:
        public Boolean VerificarRegistroBPM(IWebDriver oDriver)
        {
            try
            {
                if (oDriver.SwitchTo().Alert().ToString().Length > 0)
                {
                    oDriver.SwitchTo().Alert().Accept();
                    return false;
                }
            }
            catch (Exception)
            {
                Esperar();
            }
            return true;
        }

        public bool RecorrerGrilla(IWebDriver _driver, string idTabla, string cabecera, string valorBuscar, ref int posicionFila)
        {
            //tabla
            IList<IWebElement> _trColeccion = _driver.FindElement(By.Id(idTabla)).FindElements(By.XPath("id('" + idTabla + "')/tbody/tr"));

            int _posicionCabecera = 0;
            foreach (IWebElement item in _trColeccion)
            {
                IList<IWebElement> _td = item.FindElements(By.XPath("td"));

                for (int j = 0; j < _td.Count; j++)
                {
                    string _Cabecera = _td[j].Text;
                    if (_Cabecera.Contains(cabecera))
                    {
                        _posicionCabecera = j;
                        break;
                    }
                    if (_posicionCabecera > 0)
                    {
                        string _valorFila = _td[_posicionCabecera].Text;

                        if (_valorFila.Contains(valorBuscar))
                        {
                            return false;
                        }
                        break;
                    }
                }

                posicionFila++;
            }
            return true;
        }

    }
}