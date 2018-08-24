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
        private static Functions _Funciones;
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
        private bool _polizaNueva = true;
        private int _estadoError;
        private int _estadoFinal;

        //private static string _numeroDniContratante = string.Empty;
        //Pendiente
        private string _numeroVehiculos = "1";
        private string _numeroAsegurados = "1";

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
                    _robot.SaveTicketNextState(_Funciones.MesaDeControl(ticket, ex.Message), _robot.GetNextStateAction(ticket).First(o => o.DestinationStateId == _estadoError).Id);
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
            //LogInfoStep(5);//id referencial msje Log "Iniciando acceso al sitio policenter"

            try
            {
                _Funciones.NavegarUrlPolicyCenter(_driverGlobal, _url);
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
                _Funciones.LoginPolicyCenter(_driverGlobal, _usuario, _contraseña);
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
                _numeroPoliza = ticket.TicketValues.FirstOrDefault(np => np.FieldId == 5).Value.ToString();

                _Funciones.BuscarPolizaPolicyCenter(_driverGlobal, _numeroPoliza);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al buscar el numero de poliza", ex);
            }
            //LogInfoStep(5);//id referencial msje Log "Finalizando busqueda de poliza"

        }

        private void ObtenerDatos(Ticket ticket)
        {
            string _idDesplegable = string.Empty;
            try
            {
                _producto = _Funciones.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_PolicyDV:Product");
                _inicioVigencia = _Funciones.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:PolicyPerEffDate_date");
                _finVigencia = _Funciones.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:PolicyPerExpirDate_date");
                _tipo = _Funciones.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AssocJobDV:Type");
                _estado = _Funciones.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AssocJobDV:state");
                _tipoVigencia = _Funciones.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_DatesDV:validityType");
                _nombreContratante = _Funciones.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_AccountDV:AccountName");
                _nombreAsegurado = _Funciones.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_PolicyDV:Name");

                try
                {
                    _idDesplegable = _Funciones.ObtenerValorElemento(_driverGlobal, "FALTA ID DEL COMBOBOX");
                }
                catch
                {
                    _idDesplegable = string.Empty;
                }

                if (!string.IsNullOrEmpty(_idDesplegable))
                {
                    IList<IWebElement> _option = _driverGlobal.FindElement(By.Name("FALTA ID DEL COMBOBOX")).FindElements(By.XPath("//option"));

                    for (int h = 1; h < _option.Count; h++)
                    {
                        RecorrerGrilla();

                        if (!_polizaNueva)
                            break;

                        IList<IWebElement> _option2 = _driverGlobal.FindElement(By.Name("FALTA ID DEL COMBOBOX")).FindElements(By.XPath("//option"));
                        _option2[h].Click();
                    }
                }
                else
                {
                    RecorrerGrilla();
                }

                //Numero Canal
                _numeroCanal = _Funciones.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_ProducerDV:PolicyInfoProducerInfoSummaryInputSet:SecondaryProducerCode").Split(' ').FirstOrDefault();

                //Agente Numero y Nombre
                string[] ArrayAgente = _Funciones.ObtenerValorElemento(_driverGlobal, "PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_ProducerDV:PolicyInfoProducerInfoSummaryInputSet:ProducerCodeOfRecord").Split(' ');
                _numeroAgente = ArrayAgente[0];

                foreach (string item in ArrayAgente)
                    _agente = string.Concat(_agente, item, " ");
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
                _numeroVehiculos,_nombreContratante,_nombreAsegurado,Convert.ToString(_polizaNueva)};

                int[] IdCampos = { eesFields.Default.producto, eesFields.Default.date_inicio_vigencia, eesFields.Default.date_fin_vigencia, eesFields.Default.agente,
                eesFields.Default.num_agente,eesFields.Default.tipo_poliza,eesFields.Default.tipo_vigencia,eesFields.Default.estado_poliza,eesFields.Default.canal,eesFields.Default.num_asegurados,
                eesFields.Default.num_vehiculos,eesFields.Default.nombre_contratante,eesFields.Default.nombre_asegurado,eesFields.Default.poliza_nueva};

                for (int i = 0; i < ValorCampos.Length; i++)
                    ticket.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = ticket.Id, FieldId = IdCampos[i], Value = ValorCampos[i] });
                

            }
            catch (Exception ex)
            {
                throw new Exception("Ocurrio un Error al grabar la informacion en el ticket", ex);
            }

            try
            {
                _robot.SaveTicketNextState(ticket,_estadoFinal);
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
                _estadoError = Convert.ToInt32(_robot.GetValueParamRobot("EstadoError").ValueParam);
                _estadoFinal = Convert.ToInt32(_robot.GetValueParamRobot("EstadoSiguiente").ValueParam);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al Obtener los parametros del robot", ex);
            }
        }
     

        private void RecorrerGrilla()
        {
            //tabla
      
            IList<IWebElement> _trColeccion = _driverGlobal.FindElement(By.Id("PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_TransactionsLV")).FindElements(By.XPath("id('PolicyFile_Summary:Policy_SummaryScreen:Policy_Summary_TransactionsLV')/tbody/tr"));

            int _posicionTd = 0;
            foreach (IWebElement item in _trColeccion)
            {
                IList<IWebElement> _td = item.FindElements(By.XPath("td"));

                for (int j = 0; j < _td.Count; j++)
                {
                    string _tipoCabecera = _td[j].Text;
                    if (_tipoCabecera.Equals("Tipo"))
                    {
                        _posicionTd = j;
                        break;
                    }
                    if (_posicionTd > 0)
                    {
                        string _tipoValor = _td[_posicionTd].Text;

                        if (_tipoValor.Equals("Renovación"))
                        {
                            _polizaNueva = false;
                            break;
                        }
                        break;
                    }
                }
                if (!_polizaNueva)
                    break;
            }
        }

    }
}
