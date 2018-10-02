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
using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using OpenQA.Selenium.Interactions;
using System.Web.Script.Serialization;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Support.UI;

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
                throw new Exception("Error de acceso al sistema OnBase: " + Ex.Message, Ex);
            }
        }

        public void AbrirSelenium(ref IWebDriver oDriver)
        {
            try
            {
                oDriver = new InternetExplorerDriver();
                oDriver.Manage().Window.Maximize();
                VerificarVentanaAlerta(oDriver);
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al iniciar navegador Internet Explorer: " + Ex.Message, Ex); }
        }

        public void NavegarUrlPolicyCenter(IWebDriver oDriver, string url)
        {
            try
            {
                oDriver.Url = url;
                oDriver.Navigate().GoToUrl("javascript:document.getElementById('overridelink').click()");
                oDriver.Manage().Window.Maximize();
                Esperar();
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al ingresar a la dirección indicada: " + Ex.Message, Ex); }
        }

        public Ticket MesaDeControl(Ticket ticket, string motivo)
        {
            if (ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.error_des) == null)
            {
                ticket.TicketValues.Add(new TicketValue
                {
                    FieldId = eesFields.Default.error_des,
                    TicketId = ticket.Id,
                    Value = motivo,
                    CreationDate = DateTime.Now,
                    ClonedValueOrder = null
                });
            }
            else
                ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.error_des).Value = motivo;
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
            try
            {
                _driver.Url = url;
                _driver.Manage().Window.Maximize();
                Esperar();
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al ingresar al sitio portal bcp: " + Ex.Message, Ex); }
        }

        public void LoginPolicyCenter(IWebDriver _driver, string usuario, string contraseña)
        {
            try
            {
                _driver.SwitchTo().Frame(_driver.FindElement(By.Id("top_frame")));
                _driver.FindElement(By.Id("Login:LoginScreen:LoginDV:username")).SendKeys(usuario);
                _driver.FindElement(By.Id("Login:LoginScreen:LoginDV:password")).SendKeys(contraseña);
                _driver.FindElement(By.Id("Login:LoginScreen:LoginDV:submit")).SendKeys(Keys.Enter);
                Esperar();
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al iniciar sesión: " + Ex.Message, Ex); }
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
            try
            {
                _driver.FindElement(By.Id("TabBar:PolicyTab_arrow")).Click();
                _driver.FindElement(By.Id("TabBar:PolicyTab:PolicyTab_PolicyRetrievalItem")).SendKeys(numeroPoliza);
                _driver.FindElement(By.Id("TabBar:PolicyTab:PolicyTab_PolicyRetrievalItem")).SendKeys(Keys.Enter);
                Esperar(10);
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al buscar póliza: " + Ex.Message, Ex); }
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

        public string GetElementValue(IWebDriver driver, By idElement)
        {
            try
            {
                return driver.FindElement(idElement).Text;
            }
            catch (Exception Ex)
            {
                throw new Exception(String.Format("Ocurrió un error al obtener el elemento \"{0}\" .", idElement), Ex);
            }
        }

        public void SeleccionarCombo(IWebDriver _driver, string idElement, string valorComparar)
        {
            IList<IWebElement> oOption = _driver.FindElement(By.Id(idElement)).FindElements(By.XPath("id('" + idElement + "')/option"));

            for (int i = 0; i < oOption.Count; i++)
            {
                if (oOption[i].Text.ToUpperInvariant().Equals(valorComparar.ToUpperInvariant()))
                {
                    oOption[i].Click();
                    Esperar(3);
                    break;
                }
            }
        }

        public string ObtenerValorDominio(Ticket ticket, int idDominio)
        {
            var container = ODataContextWrapper.GetContainer();
            try
            {
                if (ticket != null)
                    return container.DomainValues.Where(p => p.Id == idDominio).FirstOrDefault().Value.Trim().ToUpperInvariant();
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
            catch (Exception Ex) { throw new Exception("No se encontró la ruta para el archivo geckodriver.exe . " + Ex.Message, Ex); }
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
            catch (Exception Ex) { throw new Exception("Ocurrió un error al abrir navegador Firefox: " + Ex.Message, Ex); }
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
            catch (Exception Ex) { throw new Exception("Ocurrió un error al seleccionar una opción de la lista: " + Ex.Message, Ex); }
        }

        //Valida campos vacíos en TicketValues:
        public Boolean ValidarCamposVacios(Ticket oTicket, int[] oCampos)
        {
            foreach (int nCampo in oCampos)
                if (String.IsNullOrWhiteSpace(oTicket.TicketValues.FirstOrDefault(o => o.FieldId == nCampo).Value.Trim()))
                    return false;
            return true;
        }

        //Anina: Verifica si existe elemento web. Parámetro de tipo By.
        public Boolean ExisteElemento(IWebDriver oDriver, By cIdElemento, int nIntentos = 1)
        {
            for (int i = 0; i < nIntentos; i++)
            {
                try { if (oDriver.FindElement(cIdElemento).Displayed) { return true; } }
                catch (NoSuchElementException) { Esperar(); }
            }
            return false;
        }

        //Verifica si ventana de alerta existe o no:
        public Boolean VerificarVentanaAlerta(IWebDriver oDriver, bool bAceptar = true)
        {
            try
            {
                if (oDriver.SwitchTo().Alert().ToString().Length > 0)
                {
                    if (bAceptar) { oDriver.SwitchTo().Alert().Accept(); }
                    else { oDriver.SwitchTo().Alert().Dismiss(); }
                    return true;
                }
            }
            catch (NoAlertPresentException) { Esperar(); }
            return false;
        }

        public bool VerificarValorGrilla(IWebDriver _driver, string idTabla, string cabecera, string valorBuscar, ref int posicionFila)
        {
            bool _valorEncontrado = false;
            try
            {
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
                                return _valorEncontrado = true;
                            }
                            break;
                        }
                    }

                    posicionFila++;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Ocurrio un error al buscar el valor en la grilla", ex);
            }

            return _valorEncontrado;
        }

        public string ObtenerJsonGrilla(IWebDriver _driver, string idtabla)
        {
            JavaScriptSerializer sr = new JavaScriptSerializer();
            List<string> ListaValoresGrilla = new List<string>();
            try
            {
                IList<IWebElement> _trColeccion = _driver.FindElement(By.Id(idtabla)).FindElements(By.XPath("id('" + idtabla + "')/tbody/tr"));

                int _count = 0;
                foreach (IWebElement item in _trColeccion)
                {
                    IList<IWebElement> _td = item.FindElements(By.XPath("td"));
                    for (int j = 0; j < _td.Count; j++)
                    {
                        if (_count == 0)
                        {
                            _count++;
                            break;
                        }
                        if (!string.IsNullOrEmpty(_td[j].Text))
                        {
                            ListaValoresGrilla.Add(_td[j].Text);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw new Exception("Ocurrio un error al obtener la data de la grilla", ex);
            }

            return sr.Serialize(ListaValoresGrilla);
        }

        public string ObtenerJson(string cadena)
        {
            JavaScriptSerializer sr = new JavaScriptSerializer();
             
            return sr.Serialize(cadena);
        }


        public FunctionalDomains<List<DomainValue>> GetDomainValuesByParameters(Func<FunctionalDomains<List<DomainValue>>, FunctionalDomains<List<DomainValue>>> SearchDomain
                                                                                , string nameFunctionalDomain
                                                                                , string[,] parametersQueryable
                                                                                , int page = 1
                                                                                , int bufferSize = 100)
        {
            Dictionary<string, List<string>> parameters = new Dictionary<string, List<string>>();
            for (int i = 0; i < parametersQueryable.GetLength(0); i++)
            {
                parameters.Add(parametersQueryable[i, 0], new List<string> { parametersQueryable[i, 1] });
            }

            FunctionalDomains<List<DomainValue>> objSearch = FunctionalDomains<List<DomainValue>>.CreateFunctionalDomainsSearch(nameFunctionalDomain, parameters, page, bufferSize);

            FunctionalDomains<List<DomainValue>> objResult = SearchDomain(objSearch);
            return objResult;
        }

        public List<int> ObtenerValoresReprocesamiento(Ticket ticket)
        {
            List<int> valores = new List<int>();
            int _reprocesoContador = 0, _idEstadoRetorno = 0;
            if (ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.reproceso_contador) != null)
            {
                _reprocesoContador = Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.reproceso_contador).Value);
            }
            valores.Add(_reprocesoContador);

            if (ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.id_estado_retorno) != null)
            {
                _idEstadoRetorno = ticket.StateId.Value;
            }
            valores.Add(_idEstadoRetorno);

            return valores;
        }

        public void GuardarValoresReprocesamiento(Ticket ticket, int reprocesoContador, int idEstadoRetorno)
        {
            if (ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.reproceso_contador) == null)
            {
                ticket.TicketValues.Add(new TicketValue
                {
                    FieldId = eesFields.Default.reproceso_contador,
                    TicketId = ticket.Id,
                    Value = reprocesoContador.ToString(),
                    CreationDate = DateTime.Now,
                    ClonedValueOrder = null
                });
            }
            else
                ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.reproceso_contador).Value = reprocesoContador.ToString();

            if (ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.id_estado_retorno) == null)
            {
                ticket.TicketValues.Add(new TicketValue
                {
                    FieldId = eesFields.Default.id_estado_retorno,
                    TicketId = ticket.Id,
                    Value = ticket.StateId.Value.ToString(),
                    CreationDate = DateTime.Now,
                    ClonedValueOrder = null
                });
            }
            else
                ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.id_estado_retorno).Value = idEstadoRetorno.ToString();
        }

        public void BuscarDocumentoPolicyCenter(IWebDriver oDriver, string cDocumento)
        {
            try
            {
                oDriver.FindElement(By.Id("TabBar:AccountTab_arrow")).Click();
                if (cDocumento.Length == 8)
                {
                    //Buscar por DNI:
                    oDriver.FindElement(By.Id("TabBar:AccountTab:AccountTab_DniSearchItemExt")).SendKeys(cDocumento + Keys.Enter);
                }
                else if (cDocumento.Length == 11)
                {
                    //Buscar por RUC:
                    oDriver.FindElement(By.Id("TabBar:AccountTab:AccountTab_RucSearchItemExt")).SendKeys(cDocumento + Keys.Enter);
                }
                else
                {
                    //Buscar por Cuenta:
                    oDriver.FindElement(By.Id("TabBar:AccountTab:AccountTab_AccountNumberSearchItem")).SendKeys(cDocumento + Keys.Enter);
                }
                Esperar(4);
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al buscar documento: " + Ex.Message, Ex); }
        }

        public void GuardarIdPlantillaNotificacion(Ticket ticket, int nIdProceso, int nIdLinea, bool bConforme = true)
        {
            try
            {
                if (ticket.TicketValues.FirstOrDefault(tv => tv.FieldId == eesFields.Default.id_archivo_tipo_adj) == null)
                {
                    ticket.TicketValues.Add(new TicketValue
                    {
                        FieldId = eesFields.Default.reproceso_contador,
                        TicketId = ticket.Id,
                        Value = Convert.ToString(GetTemplateID(nIdProceso, nIdLinea, bConforme)),
                        CreationDate = DateTime.Now,
                        ClonedValueOrder = null
                    });
                }
                else
                    ticket.TicketValues.FirstOrDefault(tv => tv.FieldId == eesFields.Default.id_archivo_tipo_adj).Value = Convert.ToString(GetTemplateID(nIdProceso, nIdLinea, bConforme));
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al guardar plantilla de notificación: " + Ex.Message, Ex); }
        }

        private int GetTemplateID(int nProceso, int nLinea, bool bTipo)
        {
            int nIdTemplate = 0;
            if (nProceso == eesDomains.Default.AnulacionPoliza)
            {
                if (nLinea == eesDomains.Default.AP_Autos)
                {
                    if (bTipo) { nIdTemplate = eesPlantillas.Default.AP_AutosConforme; } else { nIdTemplate = eesPlantillas.Default.AP_AutosRechazo; };
                }
                else if (nLinea == eesDomains.Default.AP_RRGG)
                {
                    if (bTipo) { nIdTemplate = eesPlantillas.Default.AP_RRGGConforme; } else { nIdTemplate = eesPlantillas.Default.AP_RRGGRechazo; };
                }
                else if (nLinea == eesDomains.Default.AP_BancaAlianzas)
                {
                    if (bTipo) { nIdTemplate = eesPlantillas.Default.AP_BancaAlianzasConforme; } else { nIdTemplate = eesPlantillas.Default.AP_BancaAlianzasRechazo; };
                }
                else if (nLinea == eesDomains.Default.AP_LLPP)
                {
                    if (bTipo) { nIdTemplate = eesPlantillas.Default.AP_LLPPConforme; } else { nIdTemplate = eesPlantillas.Default.AP_LLPPRechazo; };
                }
            }
            else if (nProceso == eesDomains.Default.Rehabilitacion)
            {
                if (nLinea == eesDomains.Default.RE_BancaAlianzas)
                {
                    if (bTipo) { nIdTemplate = eesPlantillas.Default.RE_BancaAlianzasConforme; } else { nIdTemplate = eesPlantillas.Default.RE_BancaAlianzasRechazo; };
                }
                else if (nLinea == eesDomains.Default.RE_LLPP)
                {
                    if (bTipo) { nIdTemplate = eesPlantillas.Default.RE_LLPPConforme; } else { nIdTemplate = eesPlantillas.Default.RE_LLPPRechazo; };
                }
            }
            else if (nProceso == eesDomains.Default.ActualizarDatosCliente)
            {
                if (nLinea == eesDomains.Default.AC_Autos)
                {
                    if (bTipo) { nIdTemplate = eesPlantillas.Default.AC_AutosConforme; } else { nIdTemplate = eesPlantillas.Default.AC_AutosRechazo; };
                }
                else if (nLinea == eesDomains.Default.AC_RRGG)
                {
                    if (bTipo) { nIdTemplate = eesPlantillas.Default.AC_RRGGConforme; } else { nIdTemplate = eesPlantillas.Default.AC_RRGGRechazo; };
                }
                else if (nLinea == eesDomains.Default.AC_BancaAlianzas)
                {
                    if (bTipo) { nIdTemplate = eesPlantillas.Default.AC_BancaAlianzasConforme; } else { nIdTemplate = eesPlantillas.Default.AC_BancaAlianzasRechazo; };
                }
                else if (nLinea == eesDomains.Default.AC_LLPP)
                {
                    if (bTipo) { nIdTemplate = eesPlantillas.Default.AC_LLPPConforme; } else { nIdTemplate = eesPlantillas.Default.AC_LLPPRechazo; };
                }
            }
            return nIdTemplate;
        }

        public String ObtenerNOrdenTrabajo(IWebDriver oDriver, string idTabla, string nombreCabecera)
        {
            try
            {
                IList<IWebElement> trColeccion = oDriver.FindElement(By.Id(idTabla)).FindElements(By.XPath("id('" + idTabla + "')/tbody/tr"));
                int posicionCabecera = 0;
                foreach (IWebElement item in trColeccion)
                {
                    IList<IWebElement> tdColeccion = item.FindElements(By.XPath("td"));

                    for (int j = 0; j < tdColeccion.Count; j++)
                    {
                        if (tdColeccion[j].Text.Contains(nombreCabecera))
                        {
                            posicionCabecera = j;
                            break;
                        }
                        if (posicionCabecera > 0)
                        {
                            return tdColeccion[posicionCabecera].Text;
                        }
                    }
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al obtener número de orden de trabajo: " + Ex.Message, Ex); }
            return string.Empty;
        }

        public int ObtenerFilasTablaHTML(IWebDriver oDriver, string cIdTabla, string cValorFinal = "sin texto")
        {
            int nFilas = 0;
            try
            {
                IList<IWebElement> oFilas = oDriver.FindElement(By.Id(cIdTabla)).FindElements(By.XPath("id('" + cIdTabla + "')/tbody/tr"));

                foreach (IWebElement nItem in oFilas)
                {
                    IList<IWebElement> oCeldas = nItem.FindElements(By.XPath("td"));
                    if (oCeldas[1].Text == cValorFinal) { break; }
                    nFilas += 1;
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al obtener número de filas: " + Ex.Message, Ex); }
            return nFilas;
        }

        public String ObtenerCadenaDeNumeros(string cTexto)
        {
            return Regex.Replace(cTexto, @"[^\d]", "");
        }

        public String ObtenerValorCeldaTabla(IWebDriver oDriver, string cIdTabla,int nIndexFila, int nIndexCelda)
        {
            try
            {
                IList<IWebElement> oFilas = oDriver.FindElement(By.Id(cIdTabla)).FindElements(By.XPath("id('" + cIdTabla + "')/tbody/tr"));

                for (int i = 0; i < oFilas.Count; i++)
                {
                    IList<IWebElement> oCeldas = oFilas[i].FindElements(By.XPath("td"));

                    for (int j = 1; j < oCeldas.Count; j++)
                    {
                        if (i == nIndexFila)
                        {
                            return oCeldas[nIndexCelda].Text;
                        }
                    }
                }
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al obtener valor de celda: " + Ex.Message, Ex); }
            return string.Empty;
        }

        public IWebElement FindElement(IWebDriver oDriver, By by, int nTiempoEsperaSegundos = 0)
        {
            try
            {
                if(nTiempoEsperaSegundos > 0)
                {
                    WebDriverWait oEsperar = new WebDriverWait(oDriver, TimeSpan.FromSeconds(nTiempoEsperaSegundos));
                    oEsperar.IgnoreExceptionTypes(typeof(NoSuchElementException));
                    return oEsperar.Until(e =>
                    {
                        if (ExisteElemento(oDriver,by,nTiempoEsperaSegundos)) { return e.FindElement(by); }
                        else { return null; }
                    });
                }
                return oDriver.FindElement(by);
            }
            catch (NoSuchElementException Ex) { throw new Exception("Ocurrió un error al obtener elemento web: " + Ex.Message, Ex); }
        }

        public void LimpiarElementoInput(IWebDriver oDriver, By oElemento)
        {
            try
            {
                oDriver.FindElement(oElemento).SendKeys(Keys.Control + "e");
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al limpiar input: " + Ex.Message, Ex); }
        }

        public String GetDomainValue(int nParentId, int nDomainColumnaId, int nLineNumber)
        {
            var oContainer = ODataContextWrapper.GetContainer();
            List<Domain> oDominios = oContainer.Domains.Expand(a => a.DomainValues).Where(b => b.ParentId == nParentId).ToList();
            return oDominios.FirstOrDefault(c => c.Id == nDomainColumnaId).DomainValues.FirstOrDefault(c => c.LineNumber == nLineNumber).Value;
        }

        public Boolean IsFieldEdit(Ticket oTicket, int nFieldId)
        {
            if (oTicket.TicketValues.FirstOrDefault(a => a.FieldId == nFieldId) != null)
            {
                if (String.IsNullOrWhiteSpace(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == nFieldId).Value)) { return false; }
                else{ return true; }
            }
            return false;
        }
    }
}