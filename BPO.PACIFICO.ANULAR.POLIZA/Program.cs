using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BPO.PACIFICO.ANULAR.POLIZA
{
    class Program : IRobot
    {
        private static BaseRobot<Program> _robot = null;
        private static IWebDriver _driverGlobal = null;
        private static IWebElement element;


        #region ParametrosRobot
        private string _url = string.Empty;
        private string _usuario = string.Empty;
        private string _contraseña = string.Empty;
        #endregion
        #region VariablesGLoables
        private static string _producto = string.Empty;
        private static string _inicioVigencia = string.Empty;
        private static string _finVigencia = string.Empty;
        private static string _numeroAgente = string.Empty;
        private static string _agente = string.Empty;
        private static string _tipo = string.Empty;
        private static string _estado = string.Empty;
        private static string _tipoVigencia = string.Empty;
        private static string _numeroCanal = string.Empty;

        #endregion
        static void Main(string[] args)
        {
        }

        protected override void Start()
        {
        }
    }
}
