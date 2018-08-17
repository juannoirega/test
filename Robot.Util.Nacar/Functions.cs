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
        #region "Parámetros"
        private static IWebDriver _oDriver = null;
        private static IWebElement _oElement = null;
        #endregion


        //Registro en BPM:
        public void IngresarBPM(string Url, string Usuario, string Contraseña)
        {
            _oDriver = new FirefoxDriver();
            _oDriver.Url = Url;
            Esperar(2);
            VentanaWindows(Usuario, Contraseña);
            var alert = _oDriver.SwitchTo().Alert();
            alert.SetAuthenticationCredentials(Usuario, Contraseña);
            alert.Accept();
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
            if (ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tranzabilidad) == null)
            {
                ticket.TicketValues.Add(new TicketValue
                {
                    FieldId = eesFields.Default.tranzabilidad,
                    TicketId = ticket.Id,
                    Value = motivo,
                    CreationDate = DateTime.Now,
                    ClonedValueOrder = null
                });
            }
            else
                ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tranzabilidad).Value = motivo;

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

        //Método para hacer pausa en segundos:
        public void Esperar(double nTiempo = 1)
        {
            Thread.Sleep(1000 * Convert.ToInt32(nTiempo));
        }

        //Ingresar usuario y contraseña en ventana windows:
        public void VentanaWindows(string cUsuario, string cContraseña)
        {
            Keyboard.KeyPress(VirtualKeyCode.SUBTRACT);
            Keyboard.KeyPress(cUsuario);
            Keyboard.KeyPress(VirtualKeyCode.TAB);
            Keyboard.KeyPress(cContraseña);
            Keyboard.KeyPress(VirtualKeyCode.RETURN);
            Esperar(2);
        }
        public void SeleccionarCombo(IWebDriver _driver, string idElement, string valorComparar)
        {
            IList<IWebElement> _option = _driver.FindElement(By.Id(idElement)).FindElements(By.XPath("id('" + idElement + "')/option"));

            for (int i = 0; i < _option.Count; i++)
            {
                if (_option[i].Text.ToUpperInvariant().Equals(valorComparar))
                {
                    _option[i].Click();
                }
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
    }
}