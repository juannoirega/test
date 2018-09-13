using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using Robot.Util.Nacar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPO.PACIFICO.PROCESARDATOS.AC
{
    public class Program : IRobot
    {
        #region "PARÁMETROS"
        private static BaseRobot<Program> _oRobot = null;
        private static string _cEstadoActualizacion;
        private static string _cLinea = string.Empty;
        private static int _nIdMesaControl;
        private static int _nIdPantallaValidacion;
        private static int _nIdNotificacion;
        private static string _cLineaAutos = string.Empty;
        private static string _cLineaLLPP = string.Empty;
        private static string _cLineaAlianzas = string.Empty;
        private static string _cLineaRRGG = string.Empty;
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
                    _oMesaControl = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdMesaControl);
                    _cLinea = _Funciones.ObtenerValorDominio(oTicket, Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.linea).Value)).ToUpperInvariant();
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
                _cEstadoActualizacion = _oRobot.GetValueParamRobot("EstadoActualizacion").ValueParam;
                _nIdMesaControl = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoError").ValueParam);
                _nIdPantallaValidacion = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoSiguiente").ValueParam);
                _nIdNotificacion = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoNotificacion").ValueParam);
                _cLineaAutos = _oRobot.GetValueParamRobot("LineaAutos").ValueParam;
                _cLineaLLPP = _oRobot.GetValueParamRobot("LineaLLPP").ValueParam;
                _cLineaAlianzas = _oRobot.GetValueParamRobot("LineaAlianzas").ValueParam;
                _cLineaRRGG = _oRobot.GetValueParamRobot("LineaRRGG").ValueParam;
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
                                       eesFields.Default.motivo_rehabilitar, eesFields.Default.poliza_est};

            //Valida Línea de la Póliza:
            if (_cLinea == _cLineaAutos)
            {
                if (ValidarDatosPoliza(oTicketDatos))
                {
                    if (!ReglasActualizacionPolizaAutos(oTicketDatos))
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
            else if (_cLinea == _cLineaLLPP)
            {
                if (ValidarDatosPoliza(oTicketDatos))
                {
                    if (!ReglasActualizacionPolizaLLPP(oTicketDatos))
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
            else if (_cLinea == _cLineaAlianzas)
            {
                if (ValidarDatosPoliza(oTicketDatos))
                {
                    if (!ReglasActualizacionPolizaAlianzas(oTicketDatos))
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
            else if (_cLinea == _cLineaRRGG)
            {
                if (ValidarDatosPoliza(oTicketDatos))
                {
                    if (!ReglasActualizacionPolizaRRGG(oTicketDatos))
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
                                       eesFields.Default.motivo_rehabilitar, eesFields.Default.poliza_est};

            if (_Funciones.ValidarCamposVacios(oTicketDatos, oCampos))
            {
                return false;
            }
            return true;
        }

        #region "REGLAS DE VALIDACIÓN"

        private Boolean ReglasActualizacionPolizaAutos(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.poliza_est).Value.ToUpperInvariant() == _cEstadoActualizacion) //Estado: VIGENTE. 
                {
                    //REGLA: Que RUC o DNI coincidan con datos de la póliza.
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


        private Boolean ReglasActualizacionPolizaLLPP(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.poliza_est).Value.ToUpperInvariant() == _cEstadoActualizacion) //Estado: VIGENTE. 
                {
                    //REGLA: Que RUC exista en SUNAT. (Fuera de alcance).
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


        private Boolean ReglasActualizacionPolizaAlianzas(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.poliza_est).Value.ToUpperInvariant() == _cEstadoActualizacion) //Estado: VIGENTE. 
                {
                    //REGLA: Que RUC o DNI coincidan con datos de la póliza.
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


        private Boolean ReglasActualizacionPolizaRRGG(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.poliza_est).Value.ToUpperInvariant() == _cEstadoActualizacion) //Estado: VIGENTE. 
                {
                    //REGLA: Que RUC o DNI coincidan con datos de la póliza.
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
