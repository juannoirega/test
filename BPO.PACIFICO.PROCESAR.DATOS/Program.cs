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
        private static BaseRobot<Program> _oRobot = null;
        private static int _nDiasArrepentimiento;
        private static int _nDiasDesistimiento;
        private static string _cLineaPersonal;
        private static string _cCanal;
        private static string _cPolizaEmision;
        private static string _cPolizaRenovacion;
        private static string _cEstadoRehabilitacion;
        private static string _cEstadoOtrosEndosos;
        private static string _cProceso = string.Empty;
        private static string _cLinea = string.Empty;
        private static int _nIdMesaControl;
        private static int _nIdPantallaValidacion;
        private static int _nIdNotificacion;
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
                    _oPantallaValidacion = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdPantallaValidacion);
                    _oNotificacion = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdNotificacion);
                    _cProceso = _Funciones.ObtenerValorDominio(oTicket, Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_proceso).Value)).ToUpperInvariant();
                    _cLinea = _Funciones.ObtenerValorDominio(oTicket, Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_de_linea).Value)).ToUpperInvariant();
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
                _nDiasArrepentimiento = Convert.ToInt32(_oRobot.GetValueParamRobot("DiasArrepentimiento").ValueParam);
                _nDiasDesistimiento = Convert.ToInt32(_oRobot.GetValueParamRobot("DiasDesistimiento").ValueParam);
                _cLineaPersonal = _oRobot.GetValueParamRobot("LineaPersonal").ValueParam;
                _cCanal = _oRobot.GetValueParamRobot("Canal").ValueParam;
                _cPolizaEmision = _oRobot.GetValueParamRobot("Poliza_Tipo_1").ValueParam;
                _cPolizaRenovacion = _oRobot.GetValueParamRobot("Poliza_Tipo_2").ValueParam;
                _cEstadoRehabilitacion = _oRobot.GetValueParamRobot("EstadoRehabilitacion").ValueParam;
                _cEstadoOtrosEndosos = _oRobot.GetValueParamRobot("EstadoOtrosEndosos").ValueParam;
                _nIdMesaControl = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoError").ValueParam);
                _nIdPantallaValidacion = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoSiguiente").ValueParam);
                _nIdNotificacion = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoNotificacion").ValueParam);
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
            switch (_cLinea)
            {
                case "AUTOS":
                    if (SwitchAutos(oTicketDatos))
                        if (ReglasDeValidacionAutos(oTicketDatos))
                        {
                            CambiarEstadoTicket(oTicketDatos, _oPantallaValidacion);
                        }
                        else
                        {
                            //Enviar a notificación de correo:
                            CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                        }
                    break;
                case "LLPP":
                    if (SwitchLLPP(oTicketDatos))
                        if (!ReglasDeValidacionLLPP(oTicketDatos))
                        {
                            //Enviar a notificación de correo:
                            CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                        };
                    break;
                case "ALIANZAS":
                    if (SwitchAlianzas(oTicketDatos))
                        if (!ReglasDeValidacionAlianzas(oTicketDatos))
                        {
                            //Enviar a notificación de correo:
                            CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                        };
                    break;
                case "RRGG":
                    if (SwitchRRGG(oTicketDatos))
                        if (!ReglasDeValidacionRRGG(oTicketDatos))
                        {
                            //Enviar a notificación de correo:
                            CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                        };
                    break;
            }
            CambiarEstadoTicket(oTicketDatos, _oPantallaValidacion);
        }

        #region "SWITCH"
        private bool SwitchAutos(Ticket oTicketDatos)
        {
            try
            {
                switch (_cProceso)
                {
                    case "ANULACION DE POLIZA":
                        if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToUpperInvariant() == _cEstadoOtrosEndosos)
                        {
                            //Crea arreglo para los campos:
                            int[] oCampos = new int[] {eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado, eesFields.Default.fecha_hora_de_email,
                                            eesFields.Default.tipo_poliza,eesFields.Default.date_inicio_vigencia,eesFields.Default.date_fin_vigencia,
                                            eesFields.Default.estado_poliza, eesFields.Default.tipo_vigencia};

                            //Valida campos no vacíos:
                            if (_Funciones.ValidarCamposVacios(oTicketDatos, oCampos))
                            {
                                //Enviar a mesa de control: Tiene campos vacíos.
                                CambiarEstadoTicket(oTicketDatos, _oMesaControl, "El ticket no cuenta con todos los datos necesarios.");
                                return false;
                            }
                        }
                        else
                        {
                            //Enviar a notificación de correo:
                            CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                            return false;
                        }
                        break;
                    case "REHABILITACION":
                        if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToUpperInvariant() != _cEstadoRehabilitacion) //Estado: CANCELADA. 
                        {
                            int[] oCampos = new int[] {eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado,
                                                            eesFields.Default.motivo_rehabilitar};

                            if (_Funciones.ValidarCamposVacios(oTicketDatos, oCampos))
                            {
                                //Enviar a mesa de control: Tiene campos vacíos.
                                CambiarEstadoTicket(oTicketDatos, _oMesaControl, "El ticket no cuenta con todos los datos necesarios.");
                                return false;
                            }
                        }
                        else
                        {
                            //Enviar a notificación de correo:
                            CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                            return false;
                        }
                        break;
                    case "ACTUALIZAR DATOS DEL CLIENTE":
                        break;
                }
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un error al validar valores del ticket: " + Convert.ToString(oTicketDatos.Id) + " en el proceso: " + _cProceso, Ex);
            }
            return true;
        }

        private bool SwitchLLPP(Ticket oTicketDatos)
        {
            switch (_cProceso)
            {
                case "ANULACION DE POLIZA":
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToUpperInvariant() == _cEstadoOtrosEndosos)
                    {
                        //Valida campos no vacíos:
                        int[] oCampos = new int[] {eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado, eesFields.Default.fecha_hora_de_email,
                                            eesFields.Default.tipo_poliza,eesFields.Default.date_inicio_vigencia,eesFields.Default.date_fin_vigencia,
                                            eesFields.Default.estado_poliza, eesFields.Default.tipo_vigencia};

                        if (_Funciones.ValidarCamposVacios(oTicketDatos, oCampos))
                        {
                            //Enviar a mesa de control: Tiene campos vacíos.
                            CambiarEstadoTicket(oTicketDatos, _oMesaControl, "El ticket no cuenta con todos los datos necesarios.");
                            return false;
                        }
                    }
                    else
                    {
                        //Enviar a notificación de correo:
                        CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                        return false;
                    }
                    break;
                case "REHABILITACION":
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToUpperInvariant() != _cEstadoRehabilitacion) //Estado: CANCELADA. 
                    {
                        int[] oCampos = new int[] {eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado,
                                                            eesFields.Default.motivo_rehabilitar};

                        if (_Funciones.ValidarCamposVacios(oTicketDatos, oCampos))
                        {
                            //Enviar a mesa de control: Tiene campos vacíos.
                            CambiarEstadoTicket(oTicketDatos, _oMesaControl, "El ticket no cuenta con todos los datos necesarios.");
                            return false;
                        }
                    }
                    else
                    {
                        //Enviar a notificación de correo:
                        CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                        return false;
                    }
                    break;
                case "ACTUALIZAR DATOS DEL CLIENTE":
                    break;
            }
            return true;
        }

        private bool SwitchAlianzas(Ticket oTicketDatos)
        {
            switch (_cProceso)
            {
                case "ANULACION DE POLIZA":
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToUpperInvariant() == _cEstadoOtrosEndosos)
                    {
                        //Valida campos no vacíos:
                        int[] oCampos = new int[] {eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado, eesFields.Default.fecha_hora_de_email,
                                            eesFields.Default.tipo_poliza,eesFields.Default.date_inicio_vigencia,eesFields.Default.date_fin_vigencia,
                                            eesFields.Default.estado_poliza, eesFields.Default.tipo_vigencia, eesFields.Default.numero_de_dni};

                        if (_Funciones.ValidarCamposVacios(oTicketDatos, oCampos))
                        {
                            //Enviar a mesa de control: Tiene campos vacíos.
                            CambiarEstadoTicket(oTicketDatos, _oMesaControl, "El ticket no cuenta con todos los datos necesarios.");
                            return false;
                        }
                    }
                    else
                    {
                        //Enviar a notificación de correo:
                        CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                        return false;
                    }
                    break;
                case "REHABILITACION":
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToUpperInvariant() != _cEstadoRehabilitacion) //Estado: CANCELADA. 
                    {
                        int[] oCampos = new int[] { eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado};

                        if (_Funciones.ValidarCamposVacios(oTicketDatos, oCampos))
                        {
                            //Enviar a mesa de control: Tiene campos vacíos.
                            CambiarEstadoTicket(oTicketDatos, _oMesaControl, "El ticket no cuenta con todos los datos necesarios.");
                            return false;
                        }
                    }
                    else
                    {
                        //Enviar a notificación de correo:
                        CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                        return false;
                    }
                    break;
                case "ACTUALIZAR DATOS DEL CLIENTE":
                    break;
            }
            return true;
        }

        private bool SwitchRRGG(Ticket oTicketDatos)
        {
            switch (_cProceso)
            {
                case "ANULACION DE POLIZA":
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToUpperInvariant() == _cEstadoOtrosEndosos)
                    {
                        //Crea arreglo para los campos:
                        int[] oCampos = new int[] {eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado, eesFields.Default.fecha_hora_de_email,
                                            eesFields.Default.tipo_poliza,eesFields.Default.date_inicio_vigencia,eesFields.Default.date_fin_vigencia,
                                            eesFields.Default.estado_poliza, eesFields.Default.tipo_vigencia};

                        //Valida campos no vacíos:
                        if (_Funciones.ValidarCamposVacios(oTicketDatos, oCampos))
                        {
                            //Enviar a mesa de control: Tiene campos vacíos.
                            CambiarEstadoTicket(oTicketDatos, _oMesaControl, "El ticket no cuenta con todos los datos necesarios.");
                            return false;
                        }
                    }
                    else
                    {
                        //Enviar a notificación de correo:
                        CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                        return false;
                    }
                    break;
                case "REHABILITACION":
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.estado_poliza).Value.ToUpperInvariant() != _cEstadoRehabilitacion) //Estado: CANCELADA. 
                    {
                        int[] oCampos = new int[] {eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado,
                                                            eesFields.Default.motivo_rehabilitar};

                        if (_Funciones.ValidarCamposVacios(oTicketDatos, oCampos))
                        {
                            //Enviar a mesa de control: Tiene campos vacíos.
                            CambiarEstadoTicket(oTicketDatos, _oMesaControl, "El ticket no cuenta con todos los datos necesarios.");
                            return false;
                        }
                    }
                    else
                    {
                        //Enviar a notificación de correo:
                        CambiarEstadoTicket(oTicketDatos, _oNotificacion);
                        return false;
                    }
                    break;
                case "ACTUALIZAR DATOS DEL CLIENTE":
                    break;
            }
            return true;
        }
        #endregion

        #region "REGLAS DE VALIDACIÓN"
        private Boolean ReglasDeValidacionAutos(Ticket oTicketDatos)
        {
            switch (_cProceso)
            {
                case "ANULACION DE POLIZA":
                    return ReglasValidacionAnulacion(oTicketDatos);
                case "REHABILITACION":
                    //No aplica.
                    break;
                case "ACTUALIZAR DATOS DEL CLIENTE":
                    break;
            }
            return true;
        }

        private Boolean ReglasDeValidacionLLPP(Ticket oTicketDatos)
        {
            switch (_cProceso)
            {
                case "ANULACION DE POLIZA":
                    return ReglasValidacionAnulacion(oTicketDatos);
                case "REHABILITACION":
                    //Que no tenga siniestros: Lo valida el mismo PolicyCenter.
                    break;
                case "ACTUALIZAR DATOS DEL CLIENTE":
                    break;
            }
            return true;
        }

        private Boolean ReglasDeValidacionAlianzas(Ticket oTicketDatos)
        {
            switch (_cProceso)
            {
                case "ANULACION DE POLIZA":
                    return ReglasValidacionAnulacion(oTicketDatos);
                case "REHABILITACION":
                    //Que no tenga siniestros: Lo valida el mismo PolicyCenter.
                    break;
                case "ACTUALIZAR DATOS DEL CLIENTE":
                    break;
            }
            return true;
        }

        private Boolean ReglasDeValidacionRRGG(Ticket oTicketDatos)
        {
            switch (_cProceso)
            {
                case "ANULACION DE POLIZA":
                    return ReglasValidacionAnulacion(oTicketDatos);
                case "REHABILITACION":
                    //No aplica.
                    break;
                case "ACTUALIZAR DATOS DEL CLIENTE":
                    break;
            }
            return true;
        }
        #endregion

        #region "REGLAS DE ANULACIÓN"
        private Boolean ReglasValidacionAnulacion(Ticket oTicketDatos)
        {
            try
            {
                TimeSpan nDiferencia = Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.fecha_hora_de_email).Value)
                        - Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.date_inicio_vigencia).Value);

                //VERIFICA QUE SEA EMISIÓN:
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value.ToUpperInvariant() == _cPolizaEmision)
                {
                    //Con devolución al 100%:
                    oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "false" });
                    if (nDiferencia.Days > _nDiasDesistimiento)
                    {
                        //Verifica que tenga VoBo:
                        //if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                        //{
                        //Con prorrateo:
                        oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "true" });
                        return false;
                        //}
                    }
                }
                else if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value.ToUpperInvariant() == _cPolizaRenovacion)//Es renovación:
                {
                    //Con devolución al 100%:
                    oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "false" });
                    if (nDiferencia.Days > _nDiasArrepentimiento)
                    {
                        //Verifica que tenga VoBo:
                        //if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                        //{
                        //Con prorrateo:
                        oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "true" });
                        return false;
                        //}
                    }
                }
                else //No es ni Emisión ni Renovación:
                {
                    return false;
                }
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un error al validar reglas de anulación para " + _cLinea, Ex);
            }
            return true;
        }
        #endregion
    }
}
