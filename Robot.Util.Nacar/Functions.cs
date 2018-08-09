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
            //_oDriver.Manage().Window.Maximize();
            //_oDriver.SwitchTo().Frame(_oDriver.FindElement(By.Id("NavPanelIFrame")));
            var alert = _oDriver.SwitchTo().Alert();
            alert.SetAuthenticationCredentials(Usuario, Contraseña);
            alert.Accept();
        }

        public void VentanaWindows(string cUsuario, string cContraseña)
        {
            Keyboard.KeyPress(VirtualKeyCode.SUBTRACT);
            Keyboard.KeyPress(cUsuario);
            Keyboard.KeyPress(VirtualKeyCode.TAB);
            Keyboard.KeyPress(cContraseña);
            Keyboard.KeyPress(VirtualKeyCode.RETURN);
            Esperar(2);
        }

        public static void AbrirSelenium(ref IWebDriver _driver)
        {
            _driver = new InternetExplorerDriver();
            _driver.Manage().Window.Maximize();
        }

        public static void NavegarUrl(IWebDriver _driver, string url)
        {
            _driver.Url = url;
            _driver.Navigate().GoToUrl("javascript:document.getElementById('overridelink').click()");
            _driver.Manage().Window.Maximize();
            Thread.Sleep(1000);
        }

        public static void Login(IWebDriver _driver, string usuario, string contraseña)
        {
            _driver.SwitchTo().Frame(_driver.FindElement(By.Id("top_frame")));

            _driver.FindElement(By.Id("Login:LoginScreen:LoginDV:username")).SendKeys(usuario);
            _driver.FindElement(By.Id("Login:LoginScreen:LoginDV:password")).SendKeys(contraseña);
            _driver.FindElement(By.Id("Login:LoginScreen:LoginDV:submit")).SendKeys(Keys.Enter);
            Thread.Sleep(1000);
        }

        public static void BuscarPoliza(IWebDriver _driver, string numeroPoliza)
        {
            _driver.FindElement(By.Id("TabBar:PolicyTab_arrow")).Click();
            _driver.FindElement(By.Id("TabBar:PolicyTab:PolicyTab_PolicyRetrievalItem")).SendKeys(numeroPoliza);
            _driver.FindElement(By.Id("TabBar:PolicyTab:PolicyTab_PolicyRetrievalItem")).SendKeys(Keys.Enter);
            Thread.Sleep(5000);
        }

        public static string ObtenerValorElemento(IWebDriver _driver, string idElemento)
        {
            IWebElement element = _driver.FindElement(By.Id(idElemento));
            string valorElemento = element.Text;
            return valorElemento;
        }

        //Método para hacer pausa en segundos:
        public void Esperar(double nTiempo = 1)
        {
            Thread.Sleep(1000 * Convert.ToInt32(nTiempo));
        }
    }
}