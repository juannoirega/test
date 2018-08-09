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
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
namespace BPO.PACIFCO.BUSCAR.POLIZA
{
    class Program : IRobot
    {
        private static BaseRobot<Program> _robot = null;
        private static IWebDriver _driverGlobal = null;
        private static IWebElement element;
        //private static Functions _Funciones;
        #region ParametrosRobot
        private string _url = string.Empty;
        private string _usuario = string.Empty;
        private string _contraseña = string.Empty;
        #endregion
        #region VariablesGLoables
        private string _numeroPoliza = string.Empty;
        private string _producto = string.Empty;
        private string _inicioVigencia = string.Empty;
        private string _finVigencia = string.Empty;
        private string _numeroAgente = string.Empty;
        private string _agente = string.Empty;
        private string _tipo = string.Empty;
        private string _estado = string.Empty;
        private string _tipoVigencia = string.Empty;
        private string _numeroCanal = string.Empty;
        private string _nombreContratante = string.Empty;
        private string _nombreAsegurado = string.Empty;
        //private static string _numeroDniContratante = string.Empty;
        //Pendiente
        private string _numeroVehiculos = string.Empty;
        private string _numeroAsegurados = string.Empty;

        #endregion
        static void Main(string[] args)
        {
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
                LogFailProcess(Constants.MSG_ERROR_EVENT_PROCESS_KEY, ex);
            }

            foreach (Ticket ticket in _robot.Tickets)
            {
                try
                {
                    ProcesarTicket(ticket);
                }
                catch (Exception ex)
                {
                    LogFailProcess(Constants.MSG_ERROR_EVENT_PROCESS_KEY, ex);
                    //Enviar a mesa control con mmensaje
                    //capturar imagen
                }
                finally
                {
                    if (_driverGlobal != null)
                        _driverGlobal.Quit();

                    LogEndStep(Constants.MSG_PROCESS_ENDED_KEY);
                }
            }
        }

        private void ProcesarTicket(Ticket ticket)
        {
            AbrirSelenium();
            NavegarUrl();
            Login();
            BuscarPoliza(ticket);
            ObtenerDatos(ticket);
            GrabarInformacion(ticket);
        }
        private void AbrirSelenium()
        {
            //LogInfoStep(5);//id referencial msje Log "Iniciando la carga Internet Explorer"
            try
            {
                Functions.AbrirSelenium(ref _driverGlobal);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Iniciar Internet Explorer", ex);
            }
            //LogInfoStep(6);//id referencial msje Log "Finalizando la carga Internet Explorer"

        }
        private void NavegarUrl()
        {
            //LogInfoStep(5);//id referencial msje Log "Iniciando acceso al sitio policenter"

            try
            {
                Functions.NavegarUrl(_driverGlobal, _url);
            }
            catch (Exception ex)
            {
                throw new Exception("No se pudo acceder al sitio policenter", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Finalizando acceso al sitio policenter"


        }
        private void Login()
        {
            //LogInfoStep(5);//id referencial msje Log "Iniciando login policenter"

            try
            {
                Functions.Login(_driverGlobal, _usuario, _contraseña);
            }
            catch (Exception ex)
            {
                throw new Exception("No se pudo acceder al sistema policenter", ex);

            }
            //LogInfoStep(5);//id referencial msje Log "Finalizacion login policenter"

        }
        private void BuscarPoliza(Ticket ticket)
        {
            //LogInfoStep(5);//id referencial msje Log "Iniciando busqueda de poliza"

            try
            {
                //obtener el numero  de Poliza
                _numeroPoliza = ticket.TicketValues.Where(np => np.FieldId == 5).ToString();

                Functions.BuscarPoliza(_driverGlobal, _numeroPoliza);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al buscar el numero de poliza", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Finalizando busqueda de poliza"

        }

        private void ObtenerDatos(Ticket ticket)
        {
            try
            {
                _producto = Functions.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_PolicyDV:Product");
                _inicioVigencia = Functions.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:PolicyPerEffDate_date");
                _finVigencia = Functions.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:PolicyPerExpirDate_date");
                _tipo = Functions.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AssocJobDV:Type");
                _estado = Functions.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AssocJobDV:state");
                _tipoVigencia = Functions.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:validityType");
                _nombreContratante = Functions.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AccountDV:AccountName");
                _nombreAsegurado = Functions.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_PolicyDV:Name");
                //FALTA NUMERO DE ASEGURADOS Y NUMERO DE VEHICULOS(CONFIRMAR COMO SE TRABAJARA)
                //POR AHORA SE VA A DEJAR EL TEMA DE LOS NUMEROS DE VEHICULOS Y NUMERO ASEGURADOS MANEJARLO COMO DATA ESTATICA POR MIENTRAS// CODEJAR COMENTADO CODIGO
                //VER TXT LOGICA PARA IMPLEMENTAR 
                if (_producto.Contains("Auto"))
                {
                    _driverGlobal.FindElement(By.Id("PolicyFile:PolicyFileAcceleratedMenuActions:PolicyMenuItemSet:PolicyMenuItemSet_Vehicles")).Click();

                    IWebElement table = _driverGlobal.FindElement(By.Id("PolicyFile_PersonalAuto_Vehicles:PolicyFile_PersonalAuto_VehiclesScreen:PAVehiclesPanelSet:VehiclesListDetailPanel:VehiclesLV"));

                    IList<IWebElement> tr_colencion = table.FindElements(By.XPath("id('PolicyFile_PersonalAuto_Vehicles:PolicyFile_PersonalAuto_VehiclesScreen:PAVehiclesPanelSet:VehiclesListDetailPanel:VehiclesLV')/tbody/tr"));

                    int count = 0;
                    foreach (IWebElement item in tr_colencion)
                    {
                        IList<IWebElement> td = item.FindElements(By.XPath("td"));


                        foreach (IWebElement tr in td)
                        {
                            if (count >= 2)
                            {
                                var dato = tr.Text;
                            }
                            count++;
                        }
                        count = 0;
                    }
                }

                //Numero Canal y Nombre
                element = _driverGlobal.FindElement(By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_ProducerDV:PolicyInfoProducerInfoSummaryInputSet:SecondaryProducerCode"));
                string _canalCadenaCompleta = element.GetAttribute("value");
                string[] ArrayCanal = _canalCadenaCompleta.Split(' ');

                int c = 0;
                foreach (string item in ArrayCanal)
                {
                    if (c == 0) { _numeroCanal = item; c++; }
                }

                //Agente Numero y Nombre
                element = _driverGlobal.FindElement(By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_ProducerDV:PolicyInfoProducerInfoSummaryInputSet:ProducerCodeOfRecord"));
                string _agenteCadenaCompleta = element.GetAttribute("value");
                string[] ArrayAgente = _agenteCadenaCompleta.Split(' ');
                int i = 0;
                foreach (string item in ArrayAgente)
                {
                    if (i == 0) { _numeroAgente = item; i++; }
                    else { _agente = string.Concat(_agente, item, " "); }
                }
            }
            catch (Exception ex)
            {

                throw new Exception("Error al Obtener los datos del sistema", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Se obtubieron los datos del sistema Policenter"


        }

        private void GrabarInformacion(Ticket ticket)
        {
            try
            {
                string[] ValorCampos = { _producto, _inicioVigencia, _finVigencia, _agente, _numeroAgente, _tipo, _tipoVigencia, _estado, _numeroCanal,_numeroAsegurados,
                _numeroVehiculos,_nombreContratante,_nombreAsegurado};

                int[] IdCampos = { eesFields.Default.producto, eesFields.Default.date_inicio_vigencia, eesFields.Default.date_fin_vigencia, eesFields.Default.agente,
                eesFields.Default.num_agente,eesFields.Default.tipo,eesFields.Default.tipo_vigencia,eesFields.Default.estado_poliza,eesFields.Default.canal,eesFields.Default.num_asegurados,
                eesFields.Default.num_vehiculos,eesFields.Default.nombre_contratante,eesFields.Default.nombre_asegurado};

                for (int i = 0; i < ValorCampos.Length; i++)
                {
                    ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, FieldId = IdCampos[i], Value = ValorCampos[i] });
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Ocurrio un Error al grabar la informacion en el ticket", ex);
            }

            try
            {

                List<StateAction> stateAction = _robot.GetNextStateAction(_robot.Tickets.FirstOrDefault());
                StateAction nextState = stateAction.Where(sa => sa.ActionDescription == "Datos Poliza").SingleOrDefault();

                if (_robot.SaveTicketNextState(ticket, nextState.Id))
                {
                    Console.WriteLine("El Ticket cambio al estado: {0}", nextState.ActionDescription);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Ocurrio un Error al avanzar al siguiente estado", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Se Guardo la Informacion en el ticket"

        }
        private void GetParameterRobots()
        {
            try
            {
                _url = _robot.GetValueParamRobot("URLPolyCenter").ValueParam;
                _usuario = _robot.GetValueParamRobot("UsuarioPolyCenter").ValueParam;
                _contraseña = _robot.GetValueParamRobot("PasswordPolyCenter").ValueParam;
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Obtener los parametros del robot", ex);
            }
        }
    }
}
