using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using everis.Ees.Proxy.Services.Interfaces;
using everis.Ees.Proxy.Core;
using Everis.Ees.Entities;
using System.Threading;
using Microsoft.VisualBasic;
using Robot.Util.Nacar;

namespace BPO.PACIFICO.ProcesarDatos
{
    public class Program : IRobot
    {
        #region "PARÁMETROS"
        private static BaseRobot<Program> _robot = null;
        private static int _nIdNombreContratante;
        private static int _nIdNombreAsegurado;
        private static int _nIdTipoPoliza;
        private static int _nIdFechaInicioVigencia;
        private static int _nIdFechaFinVigencia;
        private static int _nIdProducto;
        private static int _nIdVistoBueno;
        private static int _nIdEstado;
        private static int _nIdFechaHoraEmail;
        private static int _nIdCanal;
        private static int _nDiasArrepentimiento;
        private static int _nDiasDesistimiento;
        private static bool _bProrrata = false;
        private static string _cLineaPersonal;
        private static string _cCanal;
        private static string _cTipoPoliza;
        private static int _nEstadoPoliza;
        private static StateAction _oMesaControl;
        private static List<StateAction> _oAcciones;
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

            ObtenerParametros();
            LogStartStep(4);
            //Recorre uno a uno los tickets asignados al Robot:
            foreach (Ticket ticket in _robot.Tickets)
            {
                try
                {
                    _oAcciones = _robot.GetNextStateAction(ticket);
                    _oMesaControl = _oAcciones.Where(a => a.ActionId == 1).SingleOrDefault();
                    ProcesarTicket(ticket);
                }
                catch (Exception ex)
                {
                    LogFailStep(41, ex);
                }  
            }
            Environment.Exit(0);
        }

        //Obtiene valores para los parámetros del Robot desde EES:
        private void ObtenerParametros()
        {
            try
            {
                _nIdNombreContratante = eesFields.Default.nombre_contratante;
                _nIdNombreAsegurado = eesFields.Default.nombre_asegurado;
                _nIdTipoPoliza = eesFields.Default.tipo;
                _nIdFechaInicioVigencia = eesFields.Default.date_inicio_vigencia;
                _nIdFechaFinVigencia = eesFields.Default.date_fin_vigencia;
                _nIdProducto = eesFields.Default.producto;
                _nIdVistoBueno = eesFields.Default.vobo_producto;
                _nIdFechaHoraEmail = eesFields.Default.fecha_hora_de_email;
                _nIdCanal = eesFields.Default.canal;
                //Parámetros del Robot Procesamiento de Datos:
                _nDiasArrepentimiento = Convert.ToInt32(_robot.GetValueParamRobot("DiasArrepentimiento").ValueParam);
                _nDiasDesistimiento = Convert.ToInt32(_robot.GetValueParamRobot("DiasDesistimiento").ValueParam);
                _cLineaPersonal = _robot.GetValueParamRobot("LineaPersonal").ValueParam;
                _cCanal = _robot.GetValueParamRobot("Canal").ValueParam;
                _cTipoPoliza = _robot.GetValueParamRobot("TipoPoliza").ValueParam;
                _nEstadoPoliza = Convert.ToInt32(_robot.GetValueParamRobot("EstadoPoliza").ValueParam);
            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
            }
        }

        //Inicia el procesamiento de datos:
        private void ProcesarTicket(Ticket oTicketDatos)
        {
            try
            {
                //Valida que no tenga campos vacíos:
                if (!ValidarVacios(oTicketDatos))
                {
                    if (Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value) == _nEstadoPoliza) //ID DEL ESTADO VIGENTE. 
                    {
                        ValidaProducto(oTicketDatos);
                    }
                    else
                    {
                        //Enviar a mesa de control: No es estado VIGENTE.
                        EstadoSiguiente(oTicketDatos);
                    }  
                }
                else
                {
                    //Enviar a mesa de control: Tiene campos vacíos.
                    EstadoSiguiente(oTicketDatos);
                }
            }
            catch (Exception Ex)
            {
                LogFailStep(12, Ex);
            }
        }

        //Valida que los campos del TicketValues no estén vacíos:
        private bool ValidarVacios(Ticket oTicketDatos)
        {
            int [] oCampos = {eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado, eesFields.Default.fecha_hora_de_email};

            foreach (int campo in oCampos)
                if (String.IsNullOrWhiteSpace(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == campo).Value))
                    return false;         
                   
            return true;
        }

        //Envía el ticket al siguiente estado:
        private void EstadoSiguiente(Ticket oTicket)
        {
            _robot.SaveTicketNextState(oTicket, _oMesaControl.Id);
        }

        private void ValidaProducto(Ticket oTicketDatos)
        {
            //Valida reglas con los datos obtenidos:
            if (ReglasDeValidacion(oTicketDatos))
            {
                ActualizarTicket(oTicketDatos);
            }
            else
            {
                //Enviar a mesa de control:
                EstadoSiguiente(oTicketDatos);
            }
        }

        #region "REGLAS DE VALIDACIÓN"
        private bool ReglasDeValidacion(Ticket oTicketDatos)
        {
            TimeSpan nDiferencia = Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.fecha_hora_de_email).Value)
                        - Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdFechaInicioVigencia).Value);

            //Para saber si es Prorrata:
            TimeSpan nProrrata = Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdFechaFinVigencia).Value)
                                    - Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdFechaInicioVigencia).Value);

            //VERIFICA QUE SEA EMISIÓN:
            if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdTipoPoliza).Value == _cTipoPoliza)
            {
                if (nDiferencia.Days > _nDiasDesistimiento)
                {
                    //Verifica que tenga VoBo:
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdVistoBueno).Value.Length <= 0)
                    {
                        return false;
                    }
                }

                if (nProrrata.Days != 365)
                {
                    _bProrrata = true;
                }
                oTicketDatos.TicketValues.Add( new TicketValue { FieldId= 55, Id= oTicketDatos.Id, Value = "true", ClonedValueOrder = null });
            }
            else //Es renovación:
            {
                //Si es prorrata:
                if (nProrrata.Days != 365)
                {
                    _bProrrata = true;
                    if (nDiferencia.Days > _nDiasDesistimiento)
                    {
                        //Verifica que tenga VoBo:
                        if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdVistoBueno).Value.Length <= 0)
                        {
                            return false;
                        }
                    }
                }
                //Es devolución al 100%:
                else if (nDiferencia.Days > _nDiasArrepentimiento)
                {
                    //Verifica que tenga VoBo:
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdVistoBueno).Value.Length <= 0)
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        #endregion

        //Actualiza ticket con los nuevos datos para la anulación de póliza
        private void ActualizarTicket(Ticket oTicketActual)
        {

        }
    }
}
