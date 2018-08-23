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
        private static string _cEstadoRehabilitacion;
        private static string _cLinea = string.Empty;
        private static int _nIdMesaControl;
        private static int _nIdPantallaValidacion;
        private static int _nIdNotificacion;
        private static string _cRehabilitacionAutos = string.Empty;
        private static string _cRehabilitacionLLPP = string.Empty;
        private static string _cRehabilitacionAlianzas = string.Empty;
        private static string _cRehabilitacionRRGG = string.Empty;
        private static StateAction _oMesaControl;
        private static StateAction _oPantallaValidacion;
        private static StateAction _oNotificacion;
        private static Functions _Funciones;
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

            ObtenerParametros();
            LogStartStep(4);
            foreach (Ticket oTicket in _oRobot.Tickets)
            {
                try
                {
                    _cLinea = _Funciones.ObtenerValorDominio(oTicket, Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_linea).Value)).ToUpperInvariant();
                    _oMesaControl = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdMesaControl);
                    _oPantallaValidacion = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdPantallaValidacion);
                    _oNotificacion = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdNotificacion);
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

        //Obtiene valores para los parámetros del Robot desde EES:
        private void ObtenerParametros()
        {
            try
            {
                //Parámetros del Robot Procesamiento de Datos:
                _cEstadoRehabilitacion = _oRobot.GetValueParamRobot("EstadoRehabilitacion").ValueParam;                
                _nIdMesaControl = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoError").ValueParam);
                _nIdPantallaValidacion = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoSiguiente").ValueParam);
                _nIdNotificacion = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoNotificacion").ValueParam);
                _cRehabilitacionAutos = _oRobot.GetValueParamRobot("LineaAutos").ValueParam;
                _cRehabilitacionLLPP = _oRobot.GetValueParamRobot("LineaLLPP").ValueParam;
                _cRehabilitacionAlianzas = _oRobot.GetValueParamRobot("LineaAlianzas").ValueParam;
                _cRehabilitacionRRGG = _oRobot.GetValueParamRobot("LineaRRGG").ValueParam;
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
            //Valida Línea de la Póliza:
            if(_cLinea == _cRehabilitacionAutos)
            {
                if (ValidarDatosPoliza(oTicketDatos))
                {
                    if (!ReglasDeValidacionAutos(oTicketDatos))
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
            else if(_cLinea == _cRehabilitacionLLPP)
            {
                if (ValidarDatosPoliza(oTicketDatos))
                {
                    if (!ReglasDeValidacionLLPP(oTicketDatos))
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
            else if(_cLinea == _cRehabilitacionAlianzas)
            {
                if (ValidarDatosPoliza(oTicketDatos))
                {
                    if (!ReglasDeValidacionAlianzas(oTicketDatos))
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
            else if(_cLinea == _cRehabilitacionRRGG)
            {
                if (ValidarDatosPoliza(oTicketDatos))
                {
                    if (!ReglasDeValidacionRRGG(oTicketDatos))
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
            int[] oCampos = new int[] {eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado,
                                       eesFields.Default.motivo_rehabilitar, eesFields.Default.estado_poliza};

            if (_Funciones.ValidarCamposVacios(oTicketDatos, oCampos))
            {
                return false;
            }
            return true;
        }

        #region "REGLAS DE VALIDACIÓN"
        private Boolean ReglasDeValidacionAutos(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToUpperInvariant() == _cEstadoRehabilitacion) //Estado: CANCELADA. 
                {
                    //REGLA: Que no tenga siniestros se valida en el sistema.
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

        private Boolean ReglasDeValidacionLLPP(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToUpperInvariant() == _cEstadoRehabilitacion) //Estado: CANCELADA. 
                {
                    //REGLA: Que no tenga siniestros se valida en el sistema.
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

        private Boolean ReglasDeValidacionAlianzas(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToUpperInvariant() == _cEstadoRehabilitacion) //Estado: CANCELADA. 
                {
                    //REGLA: Que no tenga siniestros se valida en el sistema.
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

        private Boolean ReglasDeValidacionRRGG(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToUpperInvariant() == _cEstadoRehabilitacion) //Estado: CANCELADA. 
                {
                    //REGLA: Que no tenga siniestros se valida en el sistema.
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
