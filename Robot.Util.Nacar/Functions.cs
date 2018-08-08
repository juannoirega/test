using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using BPO.Framework.BPOCitrix;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            _oDriver.Manage().Window.Maximize();
            _oDriver.SwitchTo().Frame(_oDriver.FindElement(By.Id("NavPanelIFrame")));
            var alert = _oDriver.SwitchTo().Alert();
            alert.SetAuthenticationCredentials(Usuario, Contraseña);
            alert.Accept();
        }

        public void Pausa(double nTiempo = 1)
        {
            Thread.Sleep(1000 * Convert.ToInt32(nTiempo));
        }
    }
}
