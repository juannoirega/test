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
        private static int _nDiasArrepentimiento;
        private static int _nDiasDesistimiento;
        private static string _cConProrrata = "No";
        private static string _cLineaPersonal;
        private static string _cCanal;
        private static string _cTipoPoliza;
        private static string _cEstadoPoliza;
        private static int _nIdMesaControl;
        private static int _nIdPantallaValidacion;
        private static int _nIdNotificacion;
        private static StateAction _oMesaControl;
        private static StateAction _oPantallaValidacion;
        private static StateAction _oNotificacion;
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
                    _oMesaControl = _oAcciones.Where(a => a.ActionId == _nIdMesaControl).SingleOrDefault();
                    _oPantallaValidacion = _oAcciones.Where(b => b.ActionId == _nIdPantallaValidacion).SingleOrDefault();
                    _oNotificacion = _oAcciones.Where(c => c.ActionId == _nIdNotificacion).SingleOrDefault();
                    ProcesarTicket(ticket);
                }
                catch (Exception Ex)
                {
                    LogFailStep(12, Ex);
                }
            }
            Environment.Exit(0);
        }

        //Obtiene valores para los parámetros del Robot desde EES:
        private void ObtenerParametros()
        {
            try
            {
                //Parámetros del Robot Procesamiento de Datos:
                _nDiasArrepentimiento = Convert.ToInt32(_robot.GetValueParamRobot("DiasArrepentimiento").ValueParam);
                _nDiasDesistimiento = Convert.ToInt32(_robot.GetValueParamRobot("DiasDesistimiento").ValueParam);
                _cLineaPersonal = _robot.GetValueParamRobot("LineaPersonal").ValueParam;
                _cCanal = _robot.GetValueParamRobot("Canal").ValueParam;
                _cTipoPoliza = _robot.GetValueParamRobot("TipoPoliza").ValueParam;
                _cEstadoPoliza = _robot.GetValueParamRobot("EstadoPoliza").ValueParam;
                _nIdMesaControl = Convert.ToInt32(_robot.GetValueParamRobot("EstadoError").ValueParam);
                _nIdPantallaValidacion = Convert.ToInt32(_robot.GetValueParamRobot("EstadoSiguiente").ValueParam);
                _nIdNotificacion = Convert.ToInt32(_robot.GetValueParamRobot("EstadoNotificacion").ValueParam);
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
                //Valida campos no vacíos:
                if (!ValidarVacios(oTicketDatos))
                {
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToLower() == _cEstadoPoliza.ToLower()) //Estado: VIGENTE. 
                    {
                        ValidaProducto(oTicketDatos);
                    }
                    else
                    {
                        //Enviar a notificación de correo:
                        CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                    }
                }
                else
                {
                    //Enviar a mesa de control: Tiene campos vacíos.
                    CambiarEstadoTicket(oTicketDatos, _oMesaControl);
                }
            }
            catch (Exception Ex)
            {
                LogFailStep(17, Ex);
            }
        }

        //Valida que los campos del TicketValues no estén vacíos:
        private bool ValidarVacios(Ticket oTicketDatos)
        {
            int[] oCampos = new int[] {eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado, eesFields.Default.fecha_hora_de_email,
                                        eesFields.Default.tipo_poliza,eesFields.Default.date_inicio_vigencia,eesFields.Default.date_fin_vigencia,
                                        eesFields.Default.estado_poliza, eesFields.Default.tipo_vigencia};

            foreach (int campo in oCampos)
                if (String.IsNullOrWhiteSpace(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == campo).Value))
                    return false;
            return true;
        }

        private void ValidaProducto(Ticket oTicketDatos)
        {
            //Valida reglas con los datos obtenidos:
            if (ReglasDeValidacion(oTicketDatos))
            {
                AgregarNuevoTicketValue(oTicketDatos);
                CambiarEstadoTicket(oTicketDatos, _oPantallaValidacion);
            }
            else
            {
                //Enviar a notificación de correo:
                CambiarEstadoTicket(oTicketDatos, _oNotificacion);
            }
        }

        #region "REGLAS DE VALIDACIÓN"
        private bool ReglasDeValidacion(Ticket oTicketDatos)
        {
            TimeSpan nDiferencia = Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.fecha_hora_de_email).Value)
                        - Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.date_inicio_vigencia).Value);

            //Para saber si es Prorrata:
            TimeSpan nProrrata = Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.date_fin_vigencia).Value)
                                    - Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.date_inicio_vigencia).Value);

            //VERIFICA QUE SEA EMISIÓN:
            if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value == _cTipoPoliza)
            {
                if (nDiferencia.Days > _nDiasDesistimiento)
                {
                    //Verifica que tenga VoBo:
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                    {
                        return false;
                    }
                }

                if (nProrrata.Days != 365)
                {
                    _cConProrrata = "Si";
                }
            }
            else //Es renovación:
            {
                //Si es prorrata:
                if (nProrrata.Days != 365)
                {
                    _cConProrrata = "Si";
                    if (nDiferencia.Days > _nDiasDesistimiento)
                    {
                        //Verifica que tenga VoBo:
                        if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                        {
                            return false;
                        }
                    }
                }
                //Es devolución al 100%:
                else if (nDiferencia.Days > _nDiasArrepentimiento)
                {
                    //Verifica que tenga VoBo:
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        #endregion

        //Agrega nuevo ticket value al Ticket actual:
        private void AgregarNuevoTicketValue(Ticket oTicketDatos)
        {
            oTicketDatos.TicketValues.Add(new TicketValue { FieldId = eesFields.Default.aplica_prorrata, Id = oTicketDatos.Id, Value = _cConProrrata, ClonedValueOrder = null });
        }

        //Envía el ticket al siguiente estado:
        private void CambiarEstadoTicket(Ticket oTicket, StateAction oAccion)
        {
            //Estado = 1: Mesa de Control, Estado = 2: Notificación de Correo.
            _robot.SaveTicketNextState(oTicket, oAccion.Id);
        }
    }
}
