using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using OpenQA.Selenium;
using Robot.Util.Nacar;
using System;
using System.Linq;

namespace BPO.PACIFICO.REHABILITAR
{
    class Program : IRobot
    {
        private static BaseRobot<Program> _robot = null;
        private static IWebDriver _driverGlobal = null;
        private static IWebElement element;
        private static Functions _Funciones;

        #region ParametrosRobot
        private string _urlPolicyCenter = string.Empty;
        private string _usuarioPolicyCenter = string.Empty;
        private string _contraseñaPolicyCenter = string.Empty;
        private string _pasoRealizado = string.Empty;
        private int _estadoError;
        private int _estadoFinal;
        #endregion
        #region VariablesGLoables
        private int _reprocesoContador = 0;
        private int _idEstadoRetorno = 0;
        private int _plantillaConforme;
        private int _plantillaRechazo;
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
            if (!ValidarVacios(ticket))
            {
                _Funciones.AbrirSelenium(ref _driverGlobal);
                _Funciones.NavegarUrlPolicyCenter(_driverGlobal, _urlPolicyCenter);
                _Funciones.LoginPolicyCenter(_driverGlobal, _usuarioPolicyCenter, _contraseñaPolicyCenter);
                _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, ticket.TicketValues.FirstOrDefault(tv => tv.FieldId == eesFields.Default.poliza_nro).Value);
                RehabilitarPoliza(ticket);
                if (_reprocesoContador > 0)
                {
                    _reprocesoContador = 0;
                    _idEstadoRetorno = 0;
                    _Funciones.GuardarValoresReprocesamiento(ticket, _reprocesoContador, _idEstadoRetorno);
                }
                _robot.SaveTicketNextState(ticket, _estadoFinal);
            }
        }

        private void RehabilitarPoliza(Ticket ticket)
        {
            LogStartStep(44);
            try
            {
                _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions")).Click();
                if (_Funciones.ExisteElemento(_driverGlobal, "PolicyFile:PolicyFileMenuActions:PolicyFileMenuActions_NewWorkOrder:PolicyFileMenuActions_ReinstatePolicy", 2))
                {
                    _pasoRealizado = "Pestaña rehabilitar";
                    _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileMenuActions:PolicyFileMenuActions_NewWorkOrder:PolicyFileMenuActions_ReinstatePolicy")).Click();
                    _Funciones.Esperar(5);

                    string _descripcionMotivo = "";

                    _pasoRealizado = "Seleccionar motivo rehabilitar";
                    _Funciones.SeleccionarCombo(_driverGlobal, "ReinstatementWizard:ReinstatementWizard_ReinstatePolicyScreen:ReinstatePolicyDV:Reason", _Funciones.ObtenerValorDominio(ticket, Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.motivo_rehabilitar).Value)));
                    _Funciones.Esperar(2);

                    _pasoRealizado = "Ingresar descripcion";
                    _driverGlobal.FindElement(By.Id("ReinstatementWizard:ReinstatementWizard_ReinstatePolicyScreen:ReinstatePolicyDV:ReasonDescription")).SendKeys(_descripcionMotivo);
                    _Funciones.Esperar(2);

                    _pasoRealizado = "Boton cotizacion";
                    _driverGlobal.FindElement(By.XPath("//*[@id='ReinstatementWizard:ReinstatementWizard_ReinstatePolicyScreen:JobWizardToolbarButtonSet:QuoteOrReview']/span[2]")).Click();
                    _Funciones.Esperar(2);

                    _pasoRealizado = "Boton rehabilitar";
                    _driverGlobal.FindElement(By.Id("ReinstatementWizard:ReinstatementWizard_QuoteScreen:JobWizardToolbarButtonSet:Reinstate")).Click();
                    _Funciones.Esperar(1);
                    _driverGlobal.SwitchTo().Alert().Accept();
                    _Funciones.Esperar(5);

                    if (_Funciones.ExisteElementoXPath(_driverGlobal, "//*[@id='UWBlockProgressIssuesPopup:IssuesScreen:DetailsButton']/span[2]", 2))
                    {
                        _pasoRealizado = "Boton detalle";
                        _driverGlobal.FindElement(By.XPath("//*[@id='UWBlockProgressIssuesPopup:IssuesScreen:DetailsButton']/span[2]")).Click();
                        _Funciones.Esperar(3);

                        _pasoRealizado = "Check requiere rehabilitacion";
                        _driverGlobal.FindElement(By.Id("ReinstatementWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:1:UWIssueRowSet:_Checkbox")).Click();
                        _Funciones.Esperar(2);

                        _pasoRealizado = "Boton aprobar";
                        _driverGlobal.FindElement(By.XPath("//*[@id='ReinstatementWizard:Job_RiskAnalysisScreen:RiskAnalysisCV:RiskEvaluationPanelSet:Approve']/span[2]")).Click();
                        _Funciones.Esperar(2);

                        _pasoRealizado = "Check permitir edicion";
                        _driverGlobal.FindElement(By.Id("RiskApprovalDetailsPopup:0:IssueDetailsDV:UWApprovalLV:EditBeforeBind_true")).Click();
                        _Funciones.Esperar(3);

                        _pasoRealizado = "Boton aceptar";
                        _driverGlobal.FindElement(By.XPath("//*[@id='RiskApprovalDetailsPopup:Update']/span[2]")).Click();
                        _Funciones.Esperar(2);

                        _pasoRealizado = "Boton rehabilitar";
                        _driverGlobal.FindElement(By.Id("ReinstatementWizard:Job_RiskAnalysisScreen:JobWizardToolbarButtonSet:Reinstate")).Click();
                        _driverGlobal.SwitchTo().Alert().Accept();
                    }

                }
                else
                {
                    throw new Exception("Error no se encuentra la opcion Rehabilitar, verifique si la poliza ya fue rehabilitada");
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
                int[] oCampos = new int[] { eesFields.Default.motivo_rehabilitar };

                return _Funciones.ValidarCamposVacios(oTicketDatos, oCampos);
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al validar campos del Ticket: " + Convert.ToString(oTicketDatos.Id), Ex); }
        }
        private void GetParameterRobots()
        {
            _urlPolicyCenter = _robot.GetValueParamRobot("URLPolyCenter").ValueParam;
            _usuarioPolicyCenter = _robot.GetValueParamRobot("UsuarioPolyCenter").ValueParam;
            _contraseñaPolicyCenter = _robot.GetValueParamRobot("PasswordPolyCenter").ValueParam;
            _estadoError = Convert.ToInt32(_robot.GetValueParamRobot("EstadoError").ValueParam);
            _estadoFinal = Convert.ToInt32(_robot.GetValueParamRobot("EstadoSiguiente").ValueParam);
            _plantillaConforme = Convert.ToInt32(_robot.GetValueParamRobot("PlantillaConforme").ValueParam);
            _plantillaRechazo = Convert.ToInt32(_robot.GetValueParamRobot("PlantillaRechazo").ValueParam);
            LogEndStep(4);
        }
    }
}
