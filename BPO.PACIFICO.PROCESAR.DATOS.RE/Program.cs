using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using Robot.Util.Nacar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPO.PACIFICO.PROCESARDATOS.RE
{
    public class Program : IRobot
    {
        #region "PARÁMETROS"
        private static BaseRobot<Program> _oRobot = null;
        private static int _nDiasArrepentimiento;
        private static int _nDiasDesistimiento;
        private static string _cLinea = string.Empty;
        private static string _cLineaLLPP = string.Empty;
        private static string _cLineaAlianzas = string.Empty;
        private static string _cProceso = string.Empty;
        private static string _cTratamientoManual = string.Empty;
        private static string[] _procesos;
        private List<string> _productosAlianzas = new List<string>();
        private List<string> _tProductosAlianzas = new List<string>();
        private List<string> _tProductosLPersonales = new List<string>();
        private static StateAction _oMesaControl;
        private static StateAction _oPantallaValidacion;
        private static StateAction _oNotificacion;
        private static Functions _Funciones;
        private string msgConforme = string.Empty;
        private string msgNoConforme = string.Empty;
        private string msgObservacion = string.Empty;
        #endregion

        static void Main(string[] args)
        {
            try
            {
                _oRobot = new BaseRobot<Program>(args);
                _Funciones = new Functions();
                _oRobot.Start();
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex.Message);
            }
        }

        protected override void Start()
        {
            if (_oRobot.Tickets.Count < 1)
                return;
            IniciarParametros();
            ObtenerParametros();
            LogStartStep(4);
            foreach (Ticket oTicket in _oRobot.Tickets)
            {
                try
                {
                    _oMesaControl = _oRobot.GetNextStateAction(oTicket).First(a => a.ActionDescription == "Mesa de Control");
                    _cLinea = _Funciones.ObtenerValorDominio(oTicket, Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.idlinea).Value)).ToUpperInvariant();
                    _oPantallaValidacion = _oRobot.GetNextStateAction(oTicket).First(a => a.ActionDescription == "Validar");
                    _oNotificacion = _oRobot.GetNextStateAction(oTicket).First(a => a.ActionDescription == "Rechazar");                    
                    ProcesarTicket(oTicket);
                }
                catch (Exception Ex)
                {
                    CambiarEstadoTicket(oTicket, _oMesaControl, Ex.Message);
                    LogFailStep(12, Ex);
                    return;
                }
            }
        }
        //Inicializacion de los productos y tipos de productos en sus respectivos lugares
        protected void IniciarParametros()
        {
            _productosAlianzas.Add("Auto Modular | Cod. Canal 0210199");
            _productosAlianzas.Add("Auto Modular | Cod. Canal 0020962");
            _productosAlianzas.Add("PV02 (AX)");
            _productosAlianzas.Add("Auto Modular | Cod. Canal 1001330");
            _productosAlianzas.Add("Auto Modular | Cod. Canal 0024488");
            _productosAlianzas.Add("MI BANCO (AX)");
            _productosAlianzas.Add("PRE1 (AX)");
            _tProductosAlianzas.Add("PRE2 (AX)");
            _tProductosAlianzas.Add("PRE3");
            _tProductosLPersonales.Add("VTAR");
            _tProductosLPersonales.Add("VACC");
            _tProductosLPersonales.Add("VDES");
            _tProductosLPersonales.Add("PH01");
        }

        //Anina: Método para determinar a qué Línea pertenece la póliza.
        private void ObtieneLineaDeNegocio(Ticket oTicketDatos)
        {
            try
            {
                // FunctionalDomains<List<DomainValue>> objLineas = _Funciones.GetDomainValuesByParameters(_oRobot.SearchDomain, _cDominioLineas, new string[,] { { _cDominioLineasCol3, oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_linea).Value } });
                string motivoAnulacion = "(AX)";
                string producto = oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.producto).Value;
                string tipoProducto = oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.producto_tipo).Value;
                if (_productosAlianzas.Where(o => o == producto).FirstOrDefault() != null)
                {
                    _cLinea = "5";
                }
                if (_tProductosAlianzas.Where(o => o == tipoProducto).FirstOrDefault() != null)
                {
                    _cLinea = "5";

                }
                else if (_tProductosLPersonales.Where(o => o == tipoProducto).FirstOrDefault() != null)
                {
                    _cLinea = "6";
                }
                if (producto.Contains(motivoAnulacion))
                {
                    msgObservacion = "Se debe anular por lo acordado con el operador." + msgObservacion;
                }
                else if (tipoProducto.Contains(motivoAnulacion))
                {
                    msgObservacion = "Se debe anular por lo acordado con el operador." + msgObservacion;
                }
                //Reglas Conforme
                oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.idlinea, Value = _cLinea });

            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al obtener Línea de Negocio: " + Ex.Message, Ex); }
        }
        //Obtiene valores para los parámetros del Robot desde EES:
        private void ObtenerParametros()
        {
            try
            {
                //Parámetros del Robot Procesamiento de Datos:
                _nDiasArrepentimiento = Convert.ToInt32(_oRobot.GetValueParamRobot("reglaDiasPolRenovadaAuto").ValueParam);
                _nDiasDesistimiento = Convert.ToInt32(_oRobot.GetValueParamRobot("reglaDiasPolNuevaAuto").ValueParam);
                _cLineaLLPP = "LLPP";
                _cLineaAlianzas = "ALIANZAS";
                _procesos = _oRobot.GetValueParamRobot("reglaEstado").ValueParam.Split(',');
            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
            }
        }

        //Envía el ticket al siguiente estado:
        private void CambiarEstadoTicket(Ticket oTicket, StateAction oAccion, string cMensaje = "")
        {
            _oRobot.SaveTicketNextState(cMensaje == "" ? oTicket : _Funciones.MesaDeControl(oTicket, cMensaje), oAccion.Id);
        }


        //Inicia el procesamiento de datos:
        private void ProcesarTicket(Ticket oTicketDatos)
        {
            //Campos para Validar:
            int[] oCampos = new int[] {eesFields.Default.cuenta_nombre, eesFields.Default.asegurado_nombre,
                                       eesFields.Default.rehabilitar_motivo, eesFields.Default.poliza_est};

            //Valida Línea de la Póliza:
            if ((_cLinea == _cLineaAlianzas) || (_cLinea == _cLineaLLPP))
            {
                if (ValidarDatosPoliza(oTicketDatos))
                {
                    if (!ReglasActualizacionPolizaLLPPAlianzas(oTicketDatos))
                    {
                        //Enviar a notificación de correo:
                        CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                    }
                }
                else
                {
                    //Enviar a mesa de control: Tiene campos vacíos.
                    CambiarEstadoTicket(oTicketDatos, _oMesaControl, "El ticket " + Convert.ToString(oTicketDatos.Id) + " no cuenta con todos los datos necesarios.");
                    return;
                }
            }
            else
            {
                CambiarEstadoTicket(oTicketDatos, _oMesaControl, "El ticket " + Convert.ToString(oTicketDatos.Id) + " no pertenece a ninguna Línea de Negocio.");
            }
            CambiarEstadoTicket(oTicketDatos, _oPantallaValidacion);
        }

        private Boolean ValidarDatosPoliza(Ticket oTicketDatos)
        {
            int[] oCampos = new int[] {eesFields.Default.cuenta_nombre, eesFields.Default.asegurado_nombre,
                                       eesFields.Default.rehabilitar_motivo, eesFields.Default.poliza_est};

            if (_Funciones.ValidarCamposVacios(oTicketDatos, oCampos))
            {
                return false;
            }
            return true;
        }

        #region "REGLAS DE VALIDACIÓN"

        private Boolean ReglasActualizacionPolizaLLPPAlianzas(Ticket oTicketDatos)
        {
            try
            {
                string estadoPoliza = oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_est).Value;
                Boolean _bFlagVigencia = (_procesos.Where(o => o == estadoPoliza).FirstOrDefault() == null);

                if (_bFlagVigencia) //Estado: VIGENTE. 
                {
                    if (_bFlagVigencia)
                    {
                        msgObservacion = "Se debe mandar a mesa de control." + msgObservacion;
                    }
                    //SE VERIFICA SINIESTRO
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.siniestros).Value == null)
                    {
                        msgConforme = "Se esta cumpliendo la regla de la poliza sin siniestros. " + msgConforme;
                    }
                    else
                    {
                        msgNoConforme = "No se esta cumpliendo la regla de la poliza sin siniestros. " + msgNoConforme;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un error al validar reglas de Rehabilitación para " + _cLinea, Ex);
            }
            return true;
        }
         
        #endregion
    }
}
