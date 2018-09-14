using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using OpenQA.Selenium;
using Robot.Util.Nacar;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
        private string _rutaDocumentos = string.Empty;
        private string _numeroOrdenTrabajo = string.Empty;
        private int _plantillaConforme;
        private int _plantillaRechazo;
        private int _reprocesoContador = 0;
        private int _idEstadoRetorno = 0;
        private static string _pasoRealizado = string.Empty;
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
                    var valoresReprocesamiento = _Funciones.ObtenerValoresReprocesamiento(ticket);
                    if (valoresReprocesamiento.Count > 0) { _reprocesoContador = valoresReprocesamiento[0]; _idEstadoRetorno = valoresReprocesamiento[1]; }
                    ProcesarTicket(ticket);
                }
                catch (Exception ex)
                {
                    LogFailStep(30, ex);
                    _reprocesoContador++;
                    _Funciones.GuardarIdPlantillaNotificacion(ticket, _plantillaRechazo);
                    _Funciones.GuardarValoresReprocesamiento(ticket, _reprocesoContador, _idEstadoRetorno);
                    _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, ex.Message), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == _estadoError).Id);
                    return;
                }
                finally
                {
                    _Funciones.CerrarDriver(_driverGlobal);
                }
            }
        }
        private void ProcesarTicket(Ticket ticket)
        {
            //falta verificar cual sera el id del campo que confirmara si es portalbc o no**reemplazar por el "1"
            //_esPortalBcp = ticket.TicketValues.FirstOrDefault(tv => tv.FieldId == 1).Value.ToString() == "True" ? true : false;
            if (!ValidarVacios(ticket))
            {
                _Funciones.AbrirSelenium(ref _driverGlobal);
                NavegarUrl();
                Login();
                BuscarPoliza(ticket);
                AnularPoliza(ticket);
                GuardarPdf(ticket);
                _Funciones.GuardarIdPlantillaNotificacion(ticket, _plantillaConforme);
                GuardarInformacionTicket(ticket);
                if (_reprocesoContador > 0)
                {
                    _reprocesoContador = 0;
                    _idEstadoRetorno = 0;
                    _Funciones.GuardarValoresReprocesamiento(ticket, _reprocesoContador, _idEstadoRetorno);
                }
                _robot.SaveTicketNextState(ticket, _estadoFinal);
            }
        }
        private void AnularPoliza(Ticket ticket)
        {
            //if (_esPortalBcp)
            //    AnularPolizaPortalBcp();
            //else

            AnularPolizaPolicyCenter(ticket);
        }
        private void NavegarUrl()
        {
            //if (_esPortalBcp)
            //{

            //        LogInfoStep(5);//id referencial msje Log "Iniciando acceso al sitio policenter"
            //        _Funciones.NavegarUrlPortalBcp(_driverGlobal, _urlBcp);

            //    LogInfoStep(5);//id referencial msje Log "Finalizando acceso al sitio policenter"
            //}
            //else
            //{

            //LogInfoStep(5);//id referencial msje Log "Iniciando acceso al sitio policenter"
            _Funciones.NavegarUrlPolicyCenter(_driverGlobal, _urlPolicyCenter);
            //    LogInfoStep(5);//id referencial msje Log "Finalizando acceso al sitio policenter"
            //}

        }
        private void Login()
        {
            //if (_esPortalBcp)
            //{
            //        //LogInfoStep(5);//id referencial msje Log "Iniciando login policenter"
            //        _Funciones.LoginPortalBcp(_driverGlobal, _usuarioBcp, _contraseñaBcp);
            //    //LogInfoStep(5);//id referencial msje Log "Finalizacion login policenter"
            //}
            //else
            //{
            //LogInfoStep(5);//id referencial msje Log "Iniciando login policenter"
            _Funciones.LoginPolicyCenter(_driverGlobal, _usuarioPolicyCenter, _contraseñaPolicyCenter);
            //LogInfoStep(5);//id referencial msje Log "Finalizacion login policenter"
            //}

        }
        private void BuscarPoliza(Ticket ticket)
        {
            //if (_esPortalBcp)
            //{
            //    _Funciones.BuscarPolizaPortalBcp(_driverGlobal, ticket.TicketValues.FirstOrDefault(np => np.FieldId == eesFields.Default.numero_de_poliza).Value);
            //}
            //else
            //{
            _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, ticket.TicketValues.FirstOrDefault(np => np.FieldId == eesFields.Default.numero_de_poliza).Value);
            //}

        }
        //private void AnularPolizaPortalBcp()
        //{
        //    try
        //    {
        //        _driverGlobal.FindElement(By.XPath("//img[contains(@id,'ctl00_ContentPlaceHolder1_gvPolizas_ctl03_imgVer')]")).Click();
        //        _Funciones.Esperar(2);
        //        _driverGlobal.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlTipoMod']/option[2]")).Click();
        //        _Funciones.Esperar(3);
        //        _driverGlobal.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlMotivo_03']/option[2]")).Click();
        //        //Falta confirmar la anulacion de poliza portal bcp
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception("Error al anular la poliza en el sistema portal bcp", ex);
        //    }

        //}
        private void AnularPolizaPolicyCenter(Ticket ticket)
        {
            LogStartStep(45);
            try
            {
                _pasoRealizado = "Menu acciones";
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions")).Click();

                if (_Funciones.ExisteElemento(_driverGlobal, "PolicyFile:PolicyFileMenuActions:PolicyFileMenuActions_NewWorkOrder:PolicyFileMenuActions_CancelPolicy", 2))
                {
                    _pasoRealizado = "Menu opcion cancelar poliza";
                    _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions:PolicyFileMenuActions_NewWorkOrder:PolicyFileMenuActions_CancelPolicy")).Click();
                    _Funciones.Esperar(5);

                    string _descripcionMotivo = "SE DEJA CONSTANCIA POR EL PRESENTE ENDOSO QUE, LA POLIZA DEL RUBRO QUEDA CANCELADA, NULA Y SIN VALOR PARA TODOS SUS EFECTOS A PARTIR DEL";

                    _pasoRealizado = "Seleccionar combobox solictante";
                    _Funciones.SeleccionarCombo(_driverGlobal, "StartCancellation:StartCancellationScreen:CancelPolicyDV:Source", _Funciones.ObtenerValorDominio(ticket, Convert.ToInt32(ticket.TicketValues.FirstOrDefault(tv => tv.FieldId == eesFields.Default.anulacion_solicitante).Value)));
                    _Funciones.Esperar(2);

                    _pasoRealizado = "Seleccionar combobox motivo anulación";
                    _Funciones.SeleccionarCombo(_driverGlobal, "StartCancellation:StartCancellationScreen:CancelPolicyDV:Reason2", _Funciones.ObtenerValorDominio(ticket, Convert.ToInt32(ticket.TicketValues.FirstOrDefault(tv => tv.FieldId == eesFields.Default.poliza_anu_motivo).Value)));
                    _Funciones.Esperar(2);

                    _pasoRealizado = "Ingresar endoso anulacion";
                    _driverGlobal.FindElement(By.Id("StartCancellation:StartCancellationScreen:CancelPolicyDV:ReasonDescription")).SendKeys(string.Concat(_descripcionMotivo, " ", _Funciones.GetElementValue(_driverGlobal, "StartCancellation:StartCancellationScreen:CancelPolicyDV:CancelDate_date")));
                    _Funciones.Esperar(3);

                    _pasoRealizado = "Seleccionar combobox forma de reembolso";
                    _Funciones.SeleccionarCombo(_driverGlobal, "StartCancellation:StartCancellationScreen:CancelPolicyDV:CalcMethod", _Funciones.ObtenerValorDominio(ticket, Convert.ToInt32(ticket.TicketValues.FirstOrDefault(tv => tv.FieldId == eesFields.Default.forma_de_reembolso).Value)));
                    _Funciones.Esperar(2);

                    _pasoRealizado = "Iniciar cancelacion";
                    _driverGlobal.FindElement(By.XPath("//*[@id='StartCancellation:StartCancellationScreen:NewCancellation']/span[2]")).Click();
                    _Funciones.Esperar(1);

                    _driverGlobal.FindElement(By.XPath("//*[@id='CancellationWizard:CancellationWizard_QuoteScreen:JobWizardToolbarButtonSet:BindOptions_arrow']")).Click();
                    _Funciones.Esperar();
                    _pasoRealizado = "Opcion Cancelar Poliza";
                    _driverGlobal.FindElement(By.Id("CancellationWizard:CancellationWizard_QuoteScreen:JobWizardToolbarButtonSet:BindOptions:CancelNow")).Click();
                    _Funciones.Esperar();

                    _driverGlobal.SwitchTo().Alert().Accept();
                    _Funciones.Esperar();
                    _driverGlobal.FindElement(By.Id("JobComplete:JobCompleteScreen:JobCompleteDV:ViewPolicy")).Click();
                    _Funciones.Esperar(2);
                    _numeroOrdenTrabajo = _Funciones.GetElementValue(_driverGlobal, "JobComplete:JobCompleteScreen:Message").Trim().Split(' ').LastOrDefault().Replace(").", " ").Trim();
                    _Funciones.Esperar(2);
                }
                else
                {
                    _numeroOrdenTrabajo= _Funciones.ObtenerNOrdenTrabajo(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_TransactionsLV", "N.° de orden de trabajo");
                }
            }
            catch (Exception ex)
            {
                LogFailStep(12, ex); throw new Exception(ex.Message + " :" + _pasoRealizado, ex);
            }
        }
        private bool ValidarVacios(Ticket oTicketDatos)
        {
            try
            {
                int[] oCampos = new int[] { eesFields.Default.anulacion_solicitante, eesFields.Default.poliza_anu_motivo, eesFields.Default.forma_de_reembolso };

                return _Funciones.ValidarCamposVacios(oTicketDatos, oCampos);
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al validar campos del Ticket: " + Convert.ToString(oTicketDatos.Id), Ex); }
        }
        private void GuardarPdf(Ticket ticket)
        {
            LogStartStep(58);
            _Funciones.Esperar(2);
            if (_Funciones.ExisteElementoXPath(_driverGlobal, "//*[@id='PolicyFile:MenuLinks:PolicyFile_PolicyFile_Documents']/div", 2))
            {
                _pasoRealizado = "Herramientas Documentos";
                _driverGlobal.FindElement(By.XPath("//*[@id='PolicyFile:MenuLinks:PolicyFile_PolicyFile_Documents']/div")).Click();
                try
                {
                    _Funciones.Esperar(2);

                    int _posicionFila = -1;
                    if (_Funciones.VerificarValorGrilla(_driverGlobal, "PolicyFile_Documents:Policy_DocumentsScreen:DocumentsLV", "Nombre", "ACCOUNTHOLDER", ref _posicionFila))
                    {
                        _pasoRealizado = "Ver pdf";
                        _driverGlobal.FindElement(By.XPath("//*[@id='PolicyFile_Documents:Policy_DocumentsScreen:DocumentsLV:" + _posicionFila + ":DocumentsLV_ViewLink_link']")).Click();
                    }
                    else
                    {
                        _posicionFila = -1;
                        _pasoRealizado = "Buscar pdf";
                        _driverGlobal.FindElement(By.Id("PolicyFile_Documents:Policy_DocumentsScreen:Policy_DocumentSearchDV:SearchAndResetInputSet:SearchLinksInputSet:Search_link")).Click();
                        _Funciones.Esperar(3);
                        _Funciones.VerificarValorGrilla(_driverGlobal, "PolicyFile_Documents:Policy_DocumentsScreen:DocumentsLV", "Nombre", "ACCOUNTHOLDER", ref _posicionFila);
                        _pasoRealizado = "Ver pdf";
                        _driverGlobal.FindElement(By.XPath("//*[@id='PolicyFile_Documents:Policy_DocumentsScreen:DocumentsLV:" + _posicionFila + ":DocumentsLV_ViewLink_link']")).Click();
                    }
                    _Funciones.Esperar(2);
                    //Ruta donde se guardan los pdfs
                    _pasoRealizado = "Iniciando guardar pdf en directorio";
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
                catch (Exception ex) { LogFailStep(12, ex); throw new Exception(ex.Message + " :" + _pasoRealizado, ex); }
            }
            else { throw new Exception("No se encontro la opcion documentos"); }
        }

        private void GetParameterRobots()
        {
            _urlPolicyCenter = _robot.GetValueParamRobot("URLPolyCenter").ValueParam;
            _usuarioPolicyCenter = _robot.GetValueParamRobot("UsuarioPolyCenter").ValueParam;
            _contraseñaPolicyCenter = _robot.GetValueParamRobot("PasswordPolyCenter").ValueParam;
            //_usuarioBcp = _robot.GetValueParamRobot("UsuarioBcp").ValueParam;
            //_contraseñaBcp = _robot.GetValueParamRobot("PasswordBcp").ValueParam;
            //_urlBcp = _robot.GetValueParamRobot("URLBcp").ValueParam;
            _estadoError = Convert.ToInt32(_robot.GetValueParamRobot("EstadoError").ValueParam);
            _estadoFinal = Convert.ToInt32(_robot.GetValueParamRobot("EstadoSiguiente").ValueParam);
            _plantillaConforme = Convert.ToInt32(_robot.GetValueParamRobot("PlantillaConforme").ValueParam);
            _plantillaRechazo = Convert.ToInt32(_robot.GetValueParamRobot("PlantillaRechazo").ValueParam);
            _rutaDocumentos = _robot.GetValueParamRobot("RutaDocumentos").ValueParam;
            LogEndStep(4);
        }

        private void GuardarInformacionTicket(Ticket ticket)
        {
            ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, FieldId = eesFields.Default.num_orden_trabajo, Value = _numeroOrdenTrabajo });
        }
    }
}
