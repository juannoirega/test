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

        public void NavegarUrl(IWebDriver _driver, string url)
        {
            _driver.Url = url;
            _driver.Navigate().GoToUrl("javascript:document.getElementById('overridelink').click()");
            _driver.Manage().Window.Maximize();
            Esperar(1);
        }

        public void Login(IWebDriver _driver, string usuario, string contraseña)
        {
            _driver.SwitchTo().Frame(_driver.FindElement(By.Id("top_frame")));

            _driver.FindElement(By.Id("Login:LoginScreen:LoginDV:username")).SendKeys(usuario);
            _driver.FindElement(By.Id("Login:LoginScreen:LoginDV:password")).SendKeys(contraseña);
            _driver.FindElement(By.Id("Login:LoginScreen:LoginDV:submit")).SendKeys(Keys.Enter);
            Esperar(1);
        }

        public void BuscarPoliza(IWebDriver _driver, string numeroPoliza)
        {
            _driver.FindElement(By.Id("TabBar:PolicyTab_arrow")).Click();
            _driver.FindElement(By.Id("TabBar:PolicyTab:PolicyTab_PolicyRetrievalItem")).SendKeys(numeroPoliza);
            Esperar(1);
            _driver.FindElement(By.Id("TabBar:PolicyTab:PolicyTab_PolicyRetrievalItem")).SendKeys(Keys.Enter);
            Esperar(5);
        }

        public string ObtenerValorElemento(IWebDriver _driver, string idElemento)
        {
            return _driver.FindElement(By.Id(idElemento)).Text;
        }

        //Método para hacer pausa en segundos:
        public static void Esperar(double nTiempo = 1)
        {
            Thread.Sleep(1000 * Convert.ToInt32(nTiempo));
        }
        
        //Ingresar usuario y contraseña en ventana windows:
        public void VentanaWindows(string cUsuario, string cContraseña)
        {
            Keyboard.KeyPress(VirtualKeyCode.SUBTRACT);
            Esperar(1);
            Keyboard.KeyPress(cUsuario);
            Esperar(1);
            Keyboard.KeyPress(VirtualKeyCode.TAB);
            Esperar(1);
            Keyboard.KeyPress(cContraseña);
            Esperar(1);
            Keyboard.KeyPress(VirtualKeyCode.RETURN);
            Esperar(2);
        }

        public void SeleccionarCombo(IWebDriver _driver, string idElement, string valorComparar)
        {
            //id Compañia seguros
            IWebElement _ddlSelects = _driver.FindElement(By.Id(idElement));
            IList<IWebElement> _option = _ddlSelects.FindElements(By.XPath("id('" + idElement + "')/option"));

            for (int i = 0; i < _option.Count; i++)
            {
                string _valorCombo = _option[i].Text.ToUpperInvariant();

                if (_valorCombo.Equals(valorComparar))
                {
                    _option[i].Click();
                }
            }
        }

        public string ObtenerValorDominio(Ticket ticket, int idCampoDominio)
        {
            string descripcion = string.Empty;
            var container = ODataContextWrapper.GetContainer();
            try
            {
                string valor = string.Empty;
                if (ticket != null)
                {
                    DomainValue valorField = container.DomainValues.Where(p => p.Id == idCampoDominio).FirstOrDefault();
                    descripcion = valorField.Value.Trim().ToUpperInvariant();
                }
            }
            catch
            {
            } 
            return descripcion;
        }
    }
}