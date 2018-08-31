using everis.Ees.Proxy;
using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using Robot.Util.Nacar;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
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
        private static Functions _Funciones;
        #region ParametrosRobot
        private string _urlPolicyCenter = string.Empty;
        private string _usuarioPolicyCenter = string.Empty;
        private string _contraseñaPolicyCenter = string.Empty;
        private string _usuarioBcp = string.Empty;
        private string _contraseñaBcp = string.Empty;
        private string _urlBcp = string.Empty;
        private int _estadoError;
        private int _estadoFinal;
        #endregion
        #region VariablesGLoables
        private string _numeroPoliza = string.Empty;
        private string _rutaDocumentos = string.Empty;
        private string _numeroEndoso = string.Empty;

        //variable temporal
        //private static bool _esPortalBcp = false;

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

            LogStartStep(2);
            try
            {
                GetParameterRobots();
            }
            catch (Exception ex)
            {
                LogFailStep(30, ex);
            }

            foreach (Ticket ticket in _robot.Tickets)
            {
                try
                {
                    ProcesarTicket(ticket);
                }
                catch (Exception ex)
                {
                    LogFailStep(30, ex);
                    _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, ex.Message), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == _estadoError).Id);
                }
                finally
                {
                    if (_driverGlobal != null)
                        _driverGlobal.Quit();
                }
            }
        }
        private void ProcesarTicket(Ticket ticket)
        {
            //falta verificar cual sera el id del campo que confirmara si es portalbc o no**reemplazar por el "1"
            //_esPortalBcp = ticket.TicketValues.FirstOrDefault(tv => tv.FieldId == 1).Value.ToString() == "True" ? true : false;
            AbrirSelenium();
            NavegarUrl();
            Login();
            BuscarPoliza(ticket);
            AnularPoliza(ticket);
            GuardarPdf(ticket);
        }
        private void AnularPoliza(Ticket ticket)
        {
            //if (_esPortalBcp)
            //    AnularPolizaPortalBcp();
            //else
            AnularPolizaPolicyCenter(ticket);
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
            //if (_esPortalBcp)
            //{
            //    try
            //    {
            //        //LogInfoStep(5);//id referencial msje Log "Iniciando acceso al sitio policenter"
            //        _Funciones.NavegarUrlPortalBcp(_driverGlobal, _urlBcp);
            //    }
            //    catch (Exception ex)
            //    {
            //        throw new Exception("No se puede acceder al sitio portal bcp", ex);
            //    }
            //    //LogInfoStep(5);//id referencial msje Log "Finalizando acceso al sitio policenter"
            //}
            //else
            //{
            try
            {
                //LogInfoStep(5);//id referencial msje Log "Iniciando acceso al sitio policenter"
                _Funciones.NavegarUrlPolicyCenter(_driverGlobal, _urlPolicyCenter);
            }
            catch (Exception ex)
            {
                throw new Exception("No se puede acceder al sitio policycenter", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Finalizando acceso al sitio policenter"
            //}

        }

        private void Login()
        {
            //if (_esPortalBcp)
            //{
            //    try
            //    {
            //        //LogInfoStep(5);//id referencial msje Log "Iniciando login policenter"
            //        _Funciones.LoginPortalBcp(_driverGlobal, _usuarioBcp, _contraseñaBcp);
            //    }
            //    catch (Exception ex)
            //    {
            //        throw new Exception("No se puede acceder al sistema portal bcp", ex);
            //    }
            //    //LogInfoStep(5);//id referencial msje Log "Finalizacion login policenter"
            //}
            //else
            //{
            try
            {
                //LogInfoStep(5);//id referencial msje Log "Iniciando login policenter"
                _Funciones.LoginPolicyCenter(_driverGlobal, _usuarioPolicyCenter, _contraseñaPolicyCenter);
            }
            catch (Exception ex)
            {
                throw new Exception("No se puede acceder al sistema policycenter", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Finalizacion login policenter"
            //}

        }
        private void BuscarPoliza(Ticket ticket)
        {
            _numeroPoliza = ticket.TicketValues.FirstOrDefault(np => np.FieldId == 5).Value.ToString();

            //if (_esPortalBcp)
            //{
            //    try
            //    {
            //        //LogInfoStep(5);//id referencial msje Log "Iniciando busqueda de poliza"
            //        if (!string.IsNullOrEmpty(_numeroPoliza))
            //        {
            //            _Funciones.BuscarPolizaPortalBcp(_driverGlobal, _numeroPoliza);
            //        }
            //        //LogInfoStep(5);//id referencial msje Log "Finalizando busqueda de poliza"
            //    }
            //    catch (Exception ex)
            //    {
            //        throw new Exception("Error al buscar el numero de poliza portal bcp", ex);
            //    }
            //}
            //else
            //{
            try
            {
                //LogInfoStep(5);//id referencial msje Log "Iniciando busqueda de poliza"
                if (!string.IsNullOrEmpty(_numeroPoliza))
                {
                    _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, _numeroPoliza);
                }
                //LogInfoStep(5);//id referencial msje Log "Finalizando busqueda de poliza"
            }
            catch (Exception ex)
            {
                throw new Exception("Error al buscar el numero de poliza policycenter", ex);
            }
            //}

        }

        private void AnularPolizaPortalBcp()
        {
            try
            {
                _driverGlobal.FindElement(By.XPath("//img[contains(@id,'ctl00_ContentPlaceHolder1_gvPolizas_ctl03_imgVer')]")).Click();
                _Funciones.Esperar(2);
                _driverGlobal.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlTipoMod']/option[2]")).Click();
                _Funciones.Esperar(3);
                _driverGlobal.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlMotivo_03']/option[2]")).Click();
                //Falta confirmar la anulacion de poliza portal bcp
            }
            catch (Exception ex)
            {
                throw new Exception("Error al anular la poliza en el sistema portal bcp", ex);
            }

        }
        private void AnularPolizaPolicyCenter(Ticket ticket)
        {
            try
            {
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions")).Click();
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions:PolicyFileMenuActions_NewWorkOrder:PolicyFileMenuActions_CancelPolicy")).Click();
                _Funciones.Esperar(5);

                string _solicitanteIdElement = "StartCancellation:StartCancellationScreen:CancelPolicyDV:Source";
                string _motivoIdElement = "StartCancellation:StartCancellationScreen:CancelPolicyDV:Reason2";
                string _reembolsoIdElement = "StartCancellation:StartCancellationScreen:CancelPolicyDV:CalcMethod";
                string _descripcionMotivo = "SE DEJA CONSTANCIA POR EL PRESENTE ENDOSO QUE, LA POLIZA DEL RUBRO QUEDA CANCELADA, NULA Y SIN VALOR PARA TODOS SUS EFECTOS A PARTIR DEL";


                int _idDominioSolicitante = Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.solicitante).Value.ToString());
                int _idDominioMotivo = Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.motivo_anular).Value.ToString());
                int _idDominioReembolso = Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.forma_de_reembolso).Value.ToString());

                string _textoDominioSolicitante = _Funciones.ObtenerValorDominio(ticket, _idDominioSolicitante);
                _Funciones.SeleccionarCombo(_driverGlobal, _solicitanteIdElement, _textoDominioSolicitante.ToUpperInvariant());
                _Funciones.Esperar(2);

                string _textoDominioMotivo = _Funciones.ObtenerValorDominio(ticket, _idDominioMotivo);
                _Funciones.SeleccionarCombo(_driverGlobal, _motivoIdElement, _textoDominioMotivo.ToUpperInvariant());
                _Funciones.Esperar(2);

                string _fechaEfectivaCancelacion = _Funciones.ObtenerValorElemento(_driverGlobal, "StartCancellation:StartCancellationScreen:CancelPolicyDV:CancelDate_date");

                _driverGlobal.FindElement(By.Id("StartCancellation:StartCancellationScreen:CancelPolicyDV:ReasonDescription")).SendKeys(string.Concat(_descripcionMotivo, " ", _fechaEfectivaCancelacion));
                _Funciones.Esperar(2);

                string _textoDominioReembolso = _Funciones.ObtenerValorDominio(ticket, _idDominioReembolso);
                _Funciones.SeleccionarCombo(_driverGlobal, _reembolsoIdElement, _textoDominioReembolso.ToUpperInvariant());
                _Funciones.Esperar(2);

                _driverGlobal.FindElement(By.Id("StartCancellation:StartCancellationScreen:NewCancellation")).Click();
                _Funciones.Esperar(1);

                _driverGlobal.FindElement(By.Id("CancellationWizard:CancellationWizard_QuoteScreen:JobWizardToolbarButtonSet:BindOptions_arrow")).Click();
                _driverGlobal.FindElement(By.Id("CancellationWizard:CancellationWizard_QuoteScreen:JobWizardToolbarButtonSet:BindOptions:CancelNow")).Click();
                _Funciones.Esperar(1);

                _driverGlobal.SwitchTo().Alert().Accept();
                _driverGlobal.FindElement(By.Id("JobComplete:JobCompleteScreen:JobCompleteDV:ViewPolicy")).Click();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Anular la poliza en el sistema policycenter", ex);
            }
            try
            {
                IList<IWebElement> _trColeccion = _driverGlobal.FindElement(By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_TransactionsLV")).FindElements(By.XPath("id('PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_TransactionsLV')/tbody/tr"));

                int _posicionTd = 0;
                foreach (IWebElement item in _trColeccion)
                {
                    IList<IWebElement> _td = item.FindElements(By.XPath("td"));

                    for (int j = 0; j < _td.Count; j++)
                    {
                        string _tipoCabecera = _td[j].Text;
                        if (_tipoCabecera.Equals("N.° de orden de trabajo"))
                        {
                            _posicionTd = j;
                            break;
                        }
                        if (_posicionTd > 0)
                        {
                            _numeroEndoso = _td[_posicionTd].Text;
                        }
                    }
                }
            }
            catch (Exception ex) { throw new Exception("Error al buscar numero de orden de trabajo", ex); }

            ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, Value = _numeroEndoso, FieldId = eesFields.Default.numero_de_endoso });
        }
        private void GuardarPdf(Ticket ticket)
        {
            try
            {
                _Funciones.Esperar(2);
                _driverGlobal.FindElement(By.XPath("//*[@id='PolicyFile:MenuLinks:PolicyFile_PolicyFile_Documents']")).SendKeys(Keys.Enter);
                _Funciones.Esperar(2);

                string _idTabla = "PolicyFile_Documents:Policy_DocumentsScreen:DocumentsLV";
                string _cabecera = "Nombre";
                string _valorBuscar = "ACCOUNTHOLDER";
                int _posicionFila = -1;
                bool _resultado = _Funciones.RecorrerGrilla(_driverGlobal, _idTabla, _cabecera, _valorBuscar, ref _posicionFila);
                string _filaArchivo = _posicionFila.ToString();

                _driverGlobal.FindElement(By.XPath("//*[@id='PolicyFile_Documents:Policy_DocumentsScreen:DocumentsLV:" + _filaArchivo + ":DocumentsLV_ViewLink_link']")).Click();
                _Funciones.Esperar(2);

                //Ruta donde se guardan los pdfs
                DirectoryInfo directory = new DirectoryInfo(@"D:/Users/E05167/AppData/Local/Temp/Guidewire");

                var listaArchivos = directory.GetDirectories().OrderByDescending(f => f.LastWriteTime).First();
                var archivopdf = listaArchivos.GetFiles()[0].Name;

                //ruta pdf generado
                string origenArchivo = Path.Combine(listaArchivos.FullName, archivopdf);

                //Ruta Local Destino
                string rutaDestino = Path.Combine(_rutaDocumentos, archivopdf);

                //Cantidad Documentos
                int CantidadDocumentos = ticket.TicketValues.Where(t => t.FieldId == eesFields.Default.documentos).ToList().Count;

                File.Copy(origenArchivo, rutaDestino);

                ticket.TicketValues.Add(new TicketValue { Value = rutaDestino, ClonedValueOrder = CantidadDocumentos + 1, TicketId = ticket.Id, FieldId = eesFields.Default.documentos });
            }

            catch (Exception ex)
            {
                throw new Exception("Error al guardar el archivo pdf", ex);
            }

            try
            {
                _robot.SaveTicketNextState(ticket, _estadoFinal);
            }
            catch (Exception ex)
            {

                throw new Exception("Ocurrio un error al avanzar al siguiente estado", ex);
            }

        }
        private void GetParameterRobots()
        {
            try
            {
                _urlPolicyCenter = _robot.GetValueParamRobot("URLPolyCenter").ValueParam;
                _usuarioPolicyCenter = _robot.GetValueParamRobot("UsuarioPolyCenter").ValueParam;
                _contraseñaPolicyCenter = _robot.GetValueParamRobot("PasswordPolyCenter").ValueParam;
                //_usuarioBcp = _robot.GetValueParamRobot("UsuarioBcp").ValueParam;
                //_contraseñaBcp = _robot.GetValueParamRobot("PasswordBcp").ValueParam;
                //_urlBcp = _robot.GetValueParamRobot("URLBcp").ValueParam;

                _estadoError = Convert.ToInt32(_robot.GetValueParamRobot("EstadoErrorAP").ValueParam);
                _estadoFinal = Convert.ToInt32(_robot.GetValueParamRobot("EstadoSiguienteAP").ValueParam);

                LogEndStep(4);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Obtener los parametros del robot", ex);
            }
        }
    }
}
