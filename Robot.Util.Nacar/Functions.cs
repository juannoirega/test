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


            //Citrix.Keyboard.KeyPress(user);
            //PauseDb(pauseKeys);
            //Citrix.Keyboard.KeyPress(VirtualKeyCode.TAB);
            //PauseDb(pauseKeys);
            //Citrix.Keyboard.KeyPress(password);
        }

        public void AbrirSelenium()
        {
            _oDriver = new InternetExplorerDriver();
            _oDriver.Manage().Window.Maximize();
        }

        public void NavegarUrl(string url)
        {
            _oDriver.Url = url;
            _oDriver.Navigate().GoToUrl("javascript:document.getElementById('overridelink').click()");
            _oDriver.Manage().Window.Maximize();
        }

        public void Login(string usuario, string contraseña)
        {
            _oDriver.SwitchTo().Frame(_oDriver.FindElement(By.Id("top_frame")));

            _oDriver.FindElement(By.Id("Login:LoginScreen:LoginDV:username")).SendKeys(usuario);
            _oDriver.FindElement(By.Id("Login:LoginScreen:LoginDV:password")).SendKeys(contraseña);
            _oDriver.FindElement(By.Id("Login:LoginScreen:LoginDV:submit")).SendKeys(Keys.Enter);
            Thread.Sleep(300);
        }

        public void BuscarPoliza(string numeroPoliza)
        {
            _oDriver.FindElement(By.Id("TabBar:PolicyTab_arrow")).Click();
            _oDriver.FindElement(By.Id("TabBar:PolicyTab:PolicyTab_PolicyRetrievalItem")).SendKeys(numeroPoliza);
            _oDriver.FindElement(By.Id("TabBar:PolicyTab:PolicyTab_PolicyRetrievalItem")).SendKeys(Keys.Enter);
        }

        //Método para hacer pausa en segundos:
        public void Esperar(double nTiempo = 1)
        {
            Thread.Sleep(1000 * Convert.ToInt32(nTiempo));
        }
    }
}
