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
                _nIdNombreContratante = eesFields.Default.nombre_asegurado;
                _nIdNombreAsegurado = eesFields.Default.nombre_asegurado;
                _nIdTipoPoliza = eesFields.Default.nombre_asegurado;
                _nIdFechaInicioVigencia = eesFields.Default.nombre_asegurado;
                _nIdFechaFinVigencia = eesFields.Default.nombre_asegurado;
                _nIdProducto = eesFields.Default.nombre_asegurado;
                _nIdVistoBueno = eesFields.Default.nombre_asegurado;
                _nIdEstado = eesFields.Default.nombre_asegurado;
                _nIdFechaHoraEmail = eesFields.Default.nombre_asegurado;
                _nIdCanal = eesFields.Default.nombre_asegurado;
                //Parámetros del Robot Procesamiento de Datos:
                _nDiasArrepentimiento = Convert.ToInt32(_robot.GetValueParamRobot("DiasArrepentimiento").ValueParam);
                _nDiasDesistimiento = Convert.ToInt32(_robot.GetValueParamRobot("DiasDesistimiento").ValueParam);
                _cLineaPersonal = _robot.GetValueParamRobot("LineaPersonal").ValueParam;
                _cCanal = _robot.GetValueParamRobot("Canal").ValueParam;
                _cTipoPoliza = _robot.GetValueParamRobot("TipoPoliza").ValueParam;
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
                    if (Convert.ToInt32(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdEstado).Value) == 1) //ID DEL ESTADO VIGENTE. 
                    {
                        ValidaProducto(oTicketDatos);
                    }
                    else
                    {
                        //Enviar a mesa de control: No es estado VIGENTE.

                    }  
                }
                else
                {
                    //Enviar a mesa de control: Tiene campos vacíos.
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
            if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdEstado).Value.Length > 0 &&
                oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdTipoPoliza).Value.Length > 0 &&
                oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdFechaInicioVigencia).Value.Length > 0 &&
                oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdFechaFinVigencia).Value.Length > 0 &&
                oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdFechaHoraEmail).Value.Length > 0 &&
                oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdNombreContratante).Value.Length > 0 &&
                oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdNombreAsegurado).Value.Length > 0 && 
                oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdCanal).Value.Length > 0 &&
                oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdProducto).Value.Length > 0) 
                {
                    return false;
                }
            return true;
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

            }
        }

        #region "REGLAS DE VALIDACIÓN"
        private bool ReglasDeValidacion(Ticket oTicketDatos)
        {
            TimeSpan nDiferencia = Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == _nIdFechaHoraEmail).Value)
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
