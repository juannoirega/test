//using OpenQA.Selenium;
//using OpenQA.Selenium.Support.UI;
//using System;
//using System.Collections.Generic;
//using System.Configuration;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace BPO.PACIFCO.BUSCAR.POLIZA
//{
//    class Navigation : IDisposable
//    {
//        public string resultado = string.Empty;
//        private bool isHeadless { get; set; }
//        private bool isActiveUsers { get; set; }
//        private bool isCalcTimeWorked { get; set; }
//        private int vlrTimeOut { get; set; }
//        private IWebDriver driver { get; set; }
//        private WebDriverWait wait { get; set; }
//        public void Dispose()
//        {
//            throw new NotImplementedException();
//        }

//        private void GetConfigs()
//        {
//            //logData.InsertLogInfo("Início da captura dos parâmetros");
//            try
//            {
//                bool blnHeadless = false; //Informa se deve iniciar headless (por padrao, tem valor falso)
//                bool blnCalcTimeWorked = false;

//                try
//                {
//                    bool.TryParse(ConfigurationManager.AppSettings["runChromeHeadless"], out blnHeadless);
//                }
//                catch { /*se nao encontrou a chave, ou se tem valor nao booleano, mantem como false*/ }

//                int iTimeOut;
//                try
//                {
//                    int.TryParse(ConfigurationManager.AppSettings["timeoutValue"], out iTimeOut);
//                }
//                catch
//                {
//                    iTimeOut = 180;
//                }

//                try
//                {
//                    bool.TryParse(ConfigurationManager.AppSettings["calcTempoTrabalho"], out blnCalcTimeWorked);
//                }
//                catch (Exception ex)
//                {
//                    //logData.InsertLogError(string.Format("{0} [{1}]", ex.Message, ex.InnerException.ToString()));
//                }

//                isHeadless = blnHeadless;
//                vlrTimeOut = iTimeOut;
//                isActiveUsers = true;
//                isCalcTimeWorked = blnCalcTimeWorked;
//                //logData.InsertLogInfo("Fim da captura dos parâmetros");
//            }
//            catch (Exception ex)
//            {
//                //logData.InsertLogError(string.Format("{0} [{1}]", ex.Message, ex.InnerException.ToString()));
//                throw;
//            }
//        }
//        private void Initialize()
//        {
//            logData.InsertLogInfo("Início da inicialização do chrome");
//            try
//            {
//                ChromeOptions options = new ChromeOptions();
//                options.AddArgument("--disable-extensions");
//                options.AddArgument("--disable-extensions-file-access-check");
//                options.AddArgument("--disable-extensions-http-throttling");
//                options.AddArgument("--disable-infobars");
//                options.AddArgument("--enable-automation");
//                options.AddArgument("--start-maximized");
//                if (isHeadless)
//                {
//                    options.AddArgument("--headless");
//                    options.AddArgument("--disable-gpu");
//                }

//                options.AddUserProfilePreference("credentials_enable_service", false);
//                options.AddUserProfilePreference("profile.password_manager_enabled", false);

//                driver = new ChromeDriver(FindWebDrivePath(), options);
//                wait = new WebDriverWait(driver, new TimeSpan(0, 0, 0, vlrTimeOut));
//                logData.InsertLogInfo("Fim da inicialização do chrome");
//            }
//            catch (Exception ex)
//            {
//                logData.InsertLogError(string.Format("{0} [{1}]", ex.Message, ex.InnerException.ToString()));
//                throw;
//            }
//        }
//    }
//}
