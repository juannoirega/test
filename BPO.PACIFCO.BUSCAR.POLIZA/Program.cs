using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using OpenQA.Selenium;
using Robot.Util.Nacar;
using System;
using System.Collections.Generic;
using System.Linq;
namespace BPO.PACIFCO.BUSCAR.POLIZA
{
    class Program : IRobot
    {
        private static BaseRobot<Program> _robot = null;
        private static IWebDriver _driverGlobal = null;
        private static Functions _Funciones;
        private StateAction _estadoContact;
        //private static Functions _Funciones;
        #region ParametrosRobot
        private string _url = string.Empty;
        private string _usuario = string.Empty;
        private string _contraseña = string.Empty;
        private string _procesoPolicyCenter = string.Empty;
        private string _procesoContact = string.Empty;
        private string _procesoInicio = string.Empty;
        private string _estadoError;
        private int _dominioProceso;
        private int _idProceso;
        private int _tiempoEsperaLargo = 0;
        #endregion
        #region VariablesGLoables
        private string _producto = string.Empty;
        private string _tipoProducto = string.Empty;
        private string _polizaInicioVigencia = string.Empty;
        private string _polizaFinVigencia = string.Empty;
        private string _polizaEstado = string.Empty;
        private string _polizaTipoVigencia = string.Empty;
        private string _nombreContratante = string.Empty;
        private string _nombreAsegurado = string.Empty;
        private string _canalOrganizacion = string.Empty;
        private string _canalAgenteCodido = string.Empty;
        private string _canalAgente = string.Empty;
        private string _canalCodigo = string.Empty;
        private string _canal = string.Empty;
        private string _servicioOrganizacion = string.Empty;
        private string _servicioAgenteCodigo = string.Empty;
        private string _servicioAgente = string.Empty;
        private string _servicioCanalCodigo = string.Empty;
        private string _servicioCanal = string.Empty;
        private string _polizaFechaEmision = string.Empty;
        private string _numeroCuenta = string.Empty;
        private string _nroDni = string.Empty;
        private string _nroRuc = string.Empty;
        private int _reprocesoContador = 0;
        private int _idEstadoRetorno = 0;
        private int _idEstadoError = 0;
        //falta
        private string _anulacionMotivo = string.Empty;
        //JSON falta implementar
        private string _siniestros = string.Empty;
        private string _endosatarios = string.Empty;
        //
        private bool _polizaNueva = true;
        private bool _finProcesoContact = false;
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
                    _Funciones.GuardarValoresReprocesamiento(ticket, _reprocesoContador, _idEstadoRetorno);
                    _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, ex.Message), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == _idEstadoError).Id);
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
            bool _buscarPolicyCenter, _existeValorProducto, _buscarContactManager;
            int _idEstadoInicio;
            try
            {

                var container = ODataContextWrapper.GetContainer();
                //poner parametro
                List<Domain> dominios = container.Domains.Expand(dv => dv.DomainValues).Where(df => df.ParentId == _dominioProceso).ToList();
                //poner parametro
                int numero = dominios.FirstOrDefault(o => o.Name == "id").DomainValues.FirstOrDefault(o => o.Value == ticket.TicketValues.FirstOrDefault(f => f.FieldId == eesFields.Default.idproceso).Value).LineNumber;
                _existeValorProducto = ValidacionProducto(ticket);
                //poner parametro para las buquesdas de dominio
                _buscarPolicyCenter = ValidacionPoliCenter(dominios, numero);
                _buscarContactManager = ValidacionContactManager(dominios, numero);
                _idEstadoInicio = Convert.ToInt32(dominios.FirstOrDefault(o => o.Name == _procesoInicio).DomainValues.FirstOrDefault(o => o.LineNumber == numero).Value);
                _idEstadoError = Convert.ToInt32(dominios.FirstOrDefault(o => o.Name == _estadoError).DomainValues.FirstOrDefault(o => o.LineNumber == numero).Value);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Verificar el tipo de Proceso", ex);
            }

            if (_buscarPolicyCenter && ticket.TicketValues.FirstOrDefault(tv => tv.FieldId == eesFields.Default.poliza_nro).Value != "")
            {
                if (!_existeValorProducto)
                {
                    BuscarPolicyCenter(ticket);
                }
            }
            if (_buscarContactManager)
            {
                List<StateAction> accionesEstado = _robot.GetNextStateAction(_robot.Tickets.FirstOrDefault());
                _estadoContact = accionesEstado.Where(se => se.ActionDescription == "Avanzar").FirstOrDefault();
                _finProcesoContact = true;
            }
            if (_reprocesoContador > 0)
            {
                _reprocesoContador = 0;
                _idEstadoRetorno = 0;
                _Funciones.GuardarValoresReprocesamiento(ticket, _reprocesoContador, _idEstadoRetorno);
            }

            if (_finProcesoContact)
                _robot.SaveTicketNextState(ticket, _estadoContact.Id);
            else
                _robot.SaveTicketNextState(ticket, _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == _idEstadoInicio).Id);

        }
        private bool ValidacionPoliCenter(List<Domain> dominios, int numero)
        {

            try
            {
                if (dominios.FirstOrDefault(o => o.Name == _procesoPolicyCenter).DomainValues.FirstOrDefault(o => o.LineNumber == numero).Value == "1")
                    return true;
                else
                    return false;

            }
            catch { return false; }
        }
        private bool ValidacionContactManager(List<Domain> dominios, int numero)
        {

            try
            {
                if (dominios.FirstOrDefault(o => o.Name == _procesoContact).DomainValues.FirstOrDefault(o => o.LineNumber == numero).Value == "1")
                    return true;
                else
                    return false;
            }
            catch { return false; }
        }
        private bool ValidacionProducto(Ticket ticket)
        {
            try
            {
                if (!String.IsNullOrEmpty(ticket.TicketValues.FirstOrDefault(tv => tv.FieldId == eesFields.Default.producto).Value))
                    return true;
                else
                    return false;
            }
            catch { return false; }

        }
        private void BuscarPolicyCenter(Ticket ticket)
        {
            _Funciones.AbrirSelenium(ref _driverGlobal);
            _Funciones.NavegarUrlPolicyCenter(_driverGlobal, _url);
            _Funciones.LoginPolicyCenter(_driverGlobal, _usuario, _contraseña);
            _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, ticket.TicketValues.FirstOrDefault(np => np.FieldId == eesFields.Default.poliza_nro).Value);
            ObtenerDatos(ticket);
            GrabarInformacion(ticket);
        }
        private void ObtenerDatos(Ticket ticket)
        {
            LogStartStep(42);
            string _idDesplegable = string.Empty, banderaDnioRuc = string.Empty;

            if (!String.IsNullOrEmpty(_Funciones.FindElement(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AccountDV:ExtContactOfficialIDsLV:0:Type"), _tiempoEsperaLargo).Text))
            {
                banderaDnioRuc = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AccountDV:ExtContactOfficialIDsLV:0:Type"));
                if (banderaDnioRuc.Equals("RUC"))
                {
                    _nroRuc = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AccountDV:ExtContactOfficialIDsLV:0:IDDocumentNumber"));
                }
                if (banderaDnioRuc.Equals("DNI"))
                {
                    _nroDni = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AccountDV:ExtContactOfficialIDsLV:0:IDDocumentNumber"));
                }
                _producto = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_PolicyDV:Product"));
                _tipoProducto = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_PolicyDV:PolicyTypeExt"));
                _polizaInicioVigencia = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:PolicyPerEffDate_date"));
                _polizaFinVigencia = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:PolicyPerExpirDate_date"));
                _polizaEstado = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AssocJobDV:state"));
                _polizaTipoVigencia = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:validityType"));
                _nombreContratante = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AccountDV:AccountName"));
                _nombreAsegurado = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_PolicyDV:Name"));
                _numeroCuenta = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AccountDV:Number"));
                _polizaFechaEmision = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:submissionDate"));
                _canalOrganizacion = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_ProducerDV:PolicyInfoProducerInfoSummaryInputSet:POROrganization"));
                _servicioOrganizacion = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_ProducerDV:PolicyInfoProducerInfoSummaryInputSet:Producer"));
                string[] ArrayAgente = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_ProducerDV:PolicyInfoProducerInfoSummaryInputSet:SecondaryProducerCode")).Split(' ');
                if (_Funciones.ExisteElemento(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:CanceledReason")))
                {
                    _anulacionMotivo = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:CanceledReason"));
                }
                if (ArrayAgente.Length > 0)
                {
                    _canalAgenteCodido = ArrayAgente[0];

                    for (int i = 1; i < ArrayAgente.Length; i++)
                    {
                        _canalAgente = string.Concat(_canalAgente, ArrayAgente[i], " ");

                    }
                }
                string[] ArrayCanal = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_ProducerDV:PolicyInfoProducerInfoSummaryInputSet:ProducerCodeOfRecord")).Split(' ');
                if (ArrayCanal.Length > 0)
                {
                    _canalCodigo = ArrayCanal[0];
                    for (int i = 1; i < ArrayCanal.Length; i++)
                    {
                        _canal = string.Concat(_canal, ArrayCanal[i], " ");

                    }
                }

                string[] ArrayServicioAgente = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_ProducerDV:PolicyInfoProducerInfoSummaryInputSet:SecondaryProducerCodeService")).Split(' ');
                if (ArrayServicioAgente.Length > 0)
                {
                    _servicioAgenteCodigo = ArrayServicioAgente[0];

                    for (int i = 1; i < ArrayServicioAgente.Length; i++)
                    {
                        _servicioAgente = string.Concat(_servicioAgente, ArrayServicioAgente[i], " ");
                    }
                }

                string[] ArrayServicioCanal = _Funciones.GetElementValue(_driverGlobal, By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_ProducerDV:PolicyInfoProducerInfoSummaryInputSet:ProducerCode")).Split(' ');
                if (ArrayServicioCanal.Length > 0)
                {
                    _servicioCanalCodigo = ArrayServicioCanal[0];

                    for (int i = 1; i < ArrayServicioCanal.Length; i++)
                    {
                        _servicioCanal = string.Concat(_servicioCanal, ArrayServicioCanal[i], " ");
                    }
                }

                try
                {//Agregar id ComboBox Paginacion talbla 
                    _idDesplegable = _Funciones.GetElementValue(_driverGlobal, By.Id("FALTA ID DEL COMBOBOX"));
                }
                catch
                {
                    _idDesplegable = string.Empty;
                }

                int count = 0;
                if (!string.IsNullOrEmpty(_idDesplegable))
                {
                    //Agregar id ComboBox Paginacion
                    IList<IWebElement> _option = _driverGlobal.FindElement(By.Name("FALTA ID DEL COMBOBOX")).FindElements(By.XPath("//option"));

                    for (int h = 1; h < _option.Count; h++)
                    {
                        _polizaNueva = _Funciones.VerificarValorGrilla(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_TransactionsLV", "Tipo", "Renovación", ref count) == true ? false : true;

                        if (!_polizaNueva)
                            break;

                        //Agregar id ComboBox Paginacion
                        IList<IWebElement> _option2 = _driverGlobal.FindElement(By.Name("FALTA ID DEL COMBOBOX")).FindElements(By.XPath("//option"));
                        _option2[h].Click();
                    }
                }
                else
                {
                    _polizaNueva = _Funciones.VerificarValorGrilla(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_TransactionsLV", "Tipo", "Renovación", ref count) == true ? false : true;
                }

                if (_Funciones.ExisteElemento(_driverGlobal, By.XPath("//*[@id='PolicyFile:MenuLinks:PolicyFile_PolicyFile_RiskAnalysis']/div")))
                {
                    _driverGlobal.FindElement(By.XPath("//*[@id='PolicyFile:MenuLinks:PolicyFile_PolicyFile_RiskAnalysis']/div")).Click();
                    _Funciones.Esperar(3);
                    _driverGlobal.FindElement(By.Id("PolicyFile_RiskAnalysis:PolicyFile_RiskAnalysisScreen:PolicyFile_RiskAnalysisCV:PolicyFile_ClaimsCardTab")).Click();
                    _siniestros = _Funciones.ObtenerJsonGrilla(_driverGlobal, "PolicyFile_RiskAnalysis:PolicyFile_RiskAnalysisScreen:PolicyFile_RiskAnalysisCV:ClaimsLV");
                }
                if (_Funciones.ExisteElemento(_driverGlobal, By.XPath("//*[@id='PolicyFile:PolicyFileAcceleratedMenuActions:PolicyMenuItemSet:PolicyMenuItemSet_Vehicles']/div")))
                {
                    _driverGlobal.FindElement(By.XPath("//*[@id='PolicyFile:PolicyFileAcceleratedMenuActions:PolicyMenuItemSet:PolicyMenuItemSet_Vehicles']/div")).Click();
                    _Funciones.Esperar(3);
                    _driverGlobal.FindElement(By.XPath("//*[@id='PolicyFile_PersonalAuto_Vehicles:PolicyFile_PersonalAuto_VehiclesScreen:PAVehiclesPanelSet:VehiclesListDetailPanel:VehiclesDetailsCV:AdditionalInterestCardTab']")).Click();
                    _endosatarios = _Funciones.ObtenerJsonGrilla(_driverGlobal, "PolicyFile_PersonalAuto_Vehicles:PolicyFile_PersonalAuto_VehiclesScreen:PAVehiclesPanelSet:VehiclesListDetailPanel:VehiclesDetailsCV:AdditionalInterestDetailsDV:AdditionalInterestLV");
                }

            }
            else
            {
                throw new Exception("Ocurrio un error, los datos de la poliza no se cargaron");
            }


        }
        private void GrabarInformacion(Ticket ticket)
        {
            LogStartStep(43);
            try
            {
                string[] ValorCampos = { _producto, _polizaInicioVigencia, _polizaFinVigencia, _polizaTipoVigencia, _polizaEstado,_nombreContratante,_nombreAsegurado,Convert.ToString(_polizaNueva),_tipoProducto,
                _polizaFechaEmision,_canalOrganizacion,_canalAgenteCodido,_canalAgente,_canalCodigo,_canal,_servicioOrganizacion,_servicioAgenteCodigo,_servicioAgente,_servicioCanalCodigo,_servicioCanal,_numeroCuenta,_anulacionMotivo,_endosatarios,_siniestros,
                _nroDni,_nroRuc};

                int[] IdCampos = { eesFields.Default.producto, eesFields.Default.poliza_fec_ini_vig , eesFields.Default.poliza_fec_fin_vig,eesFields.Default.poliza_tipo_vig,eesFields.Default.poliza_est,
                eesFields.Default.cuenta_nombre,eesFields.Default.asegurado_nombre,eesFields.Default.flg_nuevo,eesFields.Default.producto_tipo,eesFields.Default.poliza_fec_emision,
                eesFields.Default.canal_org,eesFields.Default.canal_agente_cod,eesFields.Default.canal_agente,
                eesFields.Default.canal_cod,eesFields.Default.canal,eesFields.Default.servicio_org,eesFields.Default.servicio_agente_cod,eesFields.Default.servicio_agente,eesFields.Default.servicio_canal_cod,
                eesFields.Default.servicio_canal,eesFields.Default.cuenta_nro,eesFields.Default.poliza_anu_motivo,eesFields.Default.endosos,eesFields.Default.siniestros,eesFields.Default.poliza_nro_dni,eesFields.Default.poliza_nro_ruc};

                for (int i = 0; i < ValorCampos.Length; i++)
                    ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, FieldId = IdCampos[i], Value = ValorCampos[i] });
            }
            catch (Exception ex)
            {
                throw new Exception("Ocurrio un error al grabar los valores de la poliza en el ticket", ex);
            }

        }
        private void GetParameterRobots()
        {
            _url = _robot.GetValueParamRobot("URLPolyCenter").ValueParam;
            _usuario = _robot.GetValueParamRobot("UsuarioPolyCenter").ValueParam;
            _contraseña = _robot.GetValueParamRobot("PasswordPolyCenter").ValueParam;
            _estadoError = _robot.GetValueParamRobot("ProcesoMesaControl").ValueParam;
            _dominioProceso = Convert.ToInt32(_robot.GetValueParamRobot("DominoProcesso").ValueParam);
            _procesoPolicyCenter = _robot.GetValueParamRobot("ProcessoPolicyCenter").ValueParam;
            _procesoContact = _robot.GetValueParamRobot("ProcessoContact").ValueParam;
            _procesoInicio = _robot.GetValueParamRobot("ProcessoInicio").ValueParam;
            //Verificar como se trabajara este parametro
            _tiempoEsperaLargo = Convert.ToInt32(_robot.GetValueParamRobot("TiempoEsperaLargo").ValueParam);
            LogEndStep(4);
        }

    }
}
