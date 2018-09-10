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
using everis.Ees.Proxy.Services;

namespace BPO.PACIFICO.PROCESARDATOS.AP
{
    public class Program : IRobot
    {
        #region "PARÁMETROS"
        private static BaseRobot<Program> _oRobot = null;
        private static int _nDiasArrepentimiento;
        private static int _nDiasDesistimiento;
        private static string _cPolizaEmision;
        private static string _cPolizaRenovacion;
        private static string _cEstadoAnulacion;
        private static string _cEstadoRehabilitacion;
        private static string _cEstadoActualizacion;
        private static string _cLinea = string.Empty;
        private static int _nIdMesaControl;
        private static int _nIdPantallaValidacion;
        private static int _nIdNotificacion;
        private static string _cLineaAutos = string.Empty;
        private static string _cLineaLLPP = string.Empty;
        private static string _cLineaAlianzas = string.Empty;
        private static string _cLineaRRGG = string.Empty;
        private static bool _bEnviarNotificacion;
        private static string _cDominioLineas = string.Empty;
        private static string _cDominioProcesos = string.Empty;
        private static string _cDominioLineasCol3 = string.Empty;
        private static string _cProceso = string.Empty;
        private static string[] Procesos;
        private static StateAction _oMesaControl;
        private static StateAction _oPantallaValidacion;
        private static StateAction _oNotificacion;
        private static Functions _Funciones;
        #endregion

        static void Main(string[] args)
        {
            try
            {
                _oRobot = new BaseRobot<BPO.PACIFICO.PROCESARDATOS.AP.Program>(args);
                _Funciones = new Functions();
                _oRobot.Start();
            }
            catch (Exception Ex) { Console.WriteLine(Ex.Message); }
        }

        protected override void Start()
        {
            if (_oRobot.Tickets.Count < 1)
                return;

            ObtenerParametros();
            LogStartStep(4);
            Inicio();
            foreach (Ticket oTicket in _oRobot.Tickets)
            {
                try
                {
                    _oMesaControl = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdMesaControl);
                    _cProceso = _Funciones.ObtenerValorDominio(oTicket, Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_proceso).Value));
                    _oPantallaValidacion = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdPantallaValidacion);
                    _oNotificacion = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdNotificacion);
                    ProcesosEndoso(oTicket);
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

        private void Inicio()
        {
            Console.WriteLine("♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦ ROBOT ♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦");
            Console.WriteLine("             Robot Procesamiento de Datos              ");
            Console.WriteLine("♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦");
        }

        //Obtiene valores para los parámetros del Robot desde EES:
        private void ObtenerParametros()
        {
            try
            {
                //Parámetros del Robot Procesamiento de Datos:
                _nDiasArrepentimiento = Convert.ToInt32(_oRobot.GetValueParamRobot("DiasArrepentimiento").ValueParam);
                _nDiasDesistimiento = Convert.ToInt32(_oRobot.GetValueParamRobot("DiasDesistimiento").ValueParam);
                _cPolizaEmision = _oRobot.GetValueParamRobot("Poliza_Tipo_1").ValueParam;
                _cPolizaRenovacion = _oRobot.GetValueParamRobot("Poliza_Tipo_2").ValueParam;
                _cEstadoAnulacion = _oRobot.GetValueParamRobot("EstadoAnulacion").ValueParam;
                _cEstadoRehabilitacion = _oRobot.GetValueParamRobot("EstadoRehabilitacion").ValueParam;
                _cEstadoActualizacion = _oRobot.GetValueParamRobot("EstadoActualizacion").ValueParam;
                _nIdMesaControl = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoError").ValueParam);
                _nIdPantallaValidacion = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoSiguiente").ValueParam);
                _nIdNotificacion = Convert.ToInt32(_oRobot.GetValueParamRobot("EstadoNotificacion").ValueParam);
                _cLineaAutos = _oRobot.GetValueParamRobot("LineaAutos").ValueParam;
                _cLineaLLPP = _oRobot.GetValueParamRobot("LineaLLPP").ValueParam;
                _cLineaAlianzas = _oRobot.GetValueParamRobot("LineaAlianzas").ValueParam;
                _cLineaRRGG = _oRobot.GetValueParamRobot("LineaRRGG").ValueParam;
                _bEnviarNotificacion = Convert.ToBoolean(Convert.ToInt64(_oRobot.GetValueParamRobot("EnviarNotificacion").ValueParam));
                _cDominioLineas = _oRobot.GetValueParamRobot("DominioLineas").ValueParam;
                _cDominioProcesos = _oRobot.GetValueParamRobot("DominioProcesos").ValueParam;
                _cDominioLineasCol3 = _oRobot.GetValueParamRobot("DominioLineas_col3").ValueParam;
            }
            catch (Exception Ex) { LogFailStep(12, Ex); }
        }

        //Obtiene un arreglo con todos los procesos de endoso:
        private string[] ProcesosEndoso(Ticket oTicketDatos)
        {
            try
            {
                Procesos =  _oRobot.GetValueParamRobot("Procesos").ValueParam.Split(',');
                return Procesos;
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un erro al obtener lista de procesos de endoso: " + Ex.Message, Ex);
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
            ObtieneLineaDeNegocio(oTicketDatos);
            //Si es Anulación de Póliza:
            
           CondicionalesAnulacionPoliza(oTicketDatos);
           
        }

        //Anina: Método para determinar a qué Línea pertenece la póliza.
        private void ObtieneLineaDeNegocio(Ticket oTicketDatos)
        {
            try
            {
                FunctionalDomains<List<DomainValue>> objLineas = _Funciones.GetDomainValuesByParameters(_oRobot.SearchDomain, _cDominioLineas, new string[,] { { _cDominioLineasCol3, oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_proceso).Value} });
            }
            catch (Exception Ex) { throw new Exception("Ocurrió un error al obtener Línea de Negocio: " + Ex.Message, Ex); }
        }

        //Valida campos vacíos del ticket:
        private Boolean ValidarDatosPoliza(Ticket oTicketDatos, int[] oCampos)
        {
            if (_Funciones.ValidarCamposVacios(oTicketDatos, oCampos))
            {
                return false;
            }
            return true;
        }

        #region CONDICIONALES POR TIPO DE ENDOSO
        //Invoca a los métodos de validación de Anulación según la línea a la cual pertenece la póliza:
        private void CondicionalesAnulacionPoliza(Ticket oTicketDatos)
        {
            //Campos para Validar:
            int[] oCampos = new int[] {eesFields.Default.nombre_contratante, eesFields.Default.nombre_asegurado, eesFields.Default.fecha_hora_de_email,
                                            eesFields.Default.tipo_poliza,eesFields.Default.date_inicio_vigencia,eesFields.Default.date_fin_vigencia,
                                            eesFields.Default.poliza_est, eesFields.Default.tipo_vigencia};

            //Valida Línea de la Póliza:
            if (_cLinea == _cLineaAutos)
            {
                if (ValidarDatosPoliza(oTicketDatos, oCampos))
                {
                    if (!ReglasAnulacionPolizaAutos(oTicketDatos))
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
                if (ValidarDatosPoliza(oTicketDatos, oCampos))
                {
                    if (!ReglasAnulacionPolizaLLPP(oTicketDatos))
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
                if (ValidarDatosPoliza(oTicketDatos, oCampos))
                {
                    if (!ReglasAnulacionPolizaAlianzas(oTicketDatos))
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
                if (ValidarDatosPoliza(oTicketDatos, oCampos))
                {
                    if (!ReglasAnulacionPolizaRRGG(oTicketDatos))
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
        #endregion

        #region REGLAS DE VALIDACIÓN AUTOS
        //Reglas de validación para la línea Autos:
        private Boolean ReglasAnulacionPolizaAutos(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.admin).Value.ToUpperInvariant() == _cEstadoAnulacion) //Estado: VIGENTE. 
                {
                    TimeSpan nDiferencia = Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.fecha_hora_de_email).Value)
                                            - Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.date_inicio_vigencia).Value);

                    
                    //VERIFICA QUE SEA EMISIÓN:
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value.ToUpperInvariant() == _cPolizaEmision)
                    {
                        if (nDiferencia.Days > _nDiasDesistimiento)
                        {   
                            //Se comenta esto porque se determino que la verifiacion se veria mâs adelante 
                            //Verifica que tenga VoBo:
                            //if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                            //{
                            //Con prorrateo:
                            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "1" });
                            //}
                        }
                    }
                    else if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value.ToUpperInvariant() == _cPolizaRenovacion)//Es renovación:
                    {

                        if (nDiferencia.Days > _nDiasArrepentimiento)
                        {
                            //Se comenta esto porque se determino que la verifiacion se veria mâs adelante
                            //Verifica que tenga VoBo:
                            //if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                            //{
                            //Con prorrateo:
                            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "1" });
                            //}
                        }
                    }
                    else //No es ni Emisión ni Renovación:
                    {
                        //Con devolución al 100%: 
                        oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "0" });
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un error al validar reglas de Anulación para " + _cLinea, Ex);
            }
            return true;
        } 
        #endregion

        #region REGLAS DE VALIDACIÓN LÍNEAS PERSONALES
        //Reglas de validación para la línea Líneas Personales:
        private Boolean ReglasAnulacionPolizaLLPP(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.admin).Value.ToUpperInvariant() == _cEstadoAnulacion) //Estado: CANCELADA. 
                {
                    TimeSpan nDiferencia = Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.fecha_hora_de_email).Value)
                                            - Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.date_inicio_vigencia).Value);

                   
                    //VERIFICA QUE SEA EMISIÓN:
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value.ToUpperInvariant() == _cPolizaEmision)
                    {
                        if (nDiferencia.Days > _nDiasDesistimiento)
                        {
                            //Se comenta esto porque se determino que la verifiacion se veria mâs adelante
                            //Verifica que tenga VoBo:
                            //if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                            //{
                            //Con prorrateo:
                            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "true" });
                            //}
                        }
                    }
                    else if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value.ToUpperInvariant() == _cPolizaRenovacion)//Es renovación:
                    {

                        if (nDiferencia.Days > _nDiasArrepentimiento)
                        {
                            //Se comenta esto porque se determino que la verifiacion se veria mâs adelante
                            //Verifica que tenga VoBo:
                            //if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                            //{
                            //Con prorrateo:
                            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "true" });
                            //}
                        }
                    }
                    else //No es ni Emisión ni Renovación:
                    { 
                        //Con devolución al 100%: 
                        oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "false" });
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un error al validar reglas de Anulación para " + _cLinea, Ex);
            }
            return true;
        }
         
        #endregion

        #region REGLAS DE VALIDACIÓN BANCA Y ALIANZAS
        //Reglas de validación para la línea B&A:
        private Boolean ReglasAnulacionPolizaAlianzas(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.admin).Value.ToUpperInvariant() == _cEstadoAnulacion) //Estado: CANCELADA. 
                {
                    TimeSpan nDiferencia = Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.fecha_hora_de_email).Value)
                                            - Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.date_inicio_vigencia).Value);

                    //VERIFICA QUE SEA EMISIÓN:
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value.ToUpperInvariant() == _cPolizaEmision)
                    {
                        if (nDiferencia.Days > _nDiasDesistimiento)
                        {
                            //Se comenta esto porque se determino que la verifiacion se veria mâs adelante
                            //Verifica que tenga VoBo:
                            //if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                            //{
                            //Con prorrateo:
                            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "true" });
                            //}
                        }
                    }
                    else if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value.ToUpperInvariant() == _cPolizaRenovacion)//Es renovación:
                    {

                        if (nDiferencia.Days > _nDiasArrepentimiento)
                        {
                            //Se comenta esto porque se determino que la verifiacion se veria mâs adelante
                            //Verifica que tenga VoBo:
                            //if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                            //{
                            //Con prorrateo:
                            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "true" });
                            //}
                        }
                    }
                    else //No es ni Emisión ni Renovación:
                    {
                        //Con devolución al 100%: 
                        oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "false" });
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un error al validar reglas de Anulación para " + _cLinea, Ex);
            }
            return true;
        }

         
        #endregion

        #region REGLAS DE VALIDACIÓN RIESGOS GENERALES
        //Reglas de validación para la línea Riesgos Generales:
        private Boolean ReglasAnulacionPolizaRRGG(Ticket oTicketDatos)
        {
            try
            {
                if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.admin).Value.ToUpperInvariant() == _cEstadoAnulacion) //Estado: CANCELADA. 
                {
                    TimeSpan nDiferencia = Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.fecha_hora_de_email).Value)
                                            - Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.date_inicio_vigencia).Value);

                   
                    //VERIFICA QUE SEA EMISIÓN:
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value.ToUpperInvariant() == _cPolizaEmision)
                    {
                        if (nDiferencia.Days > _nDiasDesistimiento)
                        {
                            //Se comenta esto porque se determino que la verifiacion se veria mâs adelante
                            //Verifica que tenga VoBo:
                            //if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                            //{
                            //Con prorrateo:
                            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "true" });
                            //}
                        }
                    }
                    else if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value.ToUpperInvariant() == _cPolizaRenovacion)//Es renovación:
                    {

                        if (nDiferencia.Days > _nDiasArrepentimiento)
                        {
                            //Se comenta esto porque se determino que la verifiacion se veria mâs adelante
                            //Verifica que tenga VoBo:
                            //if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.vobo_producto).Value.Length <= 0)
                            //{
                            //Con prorrateo:
                            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "true" });
                            //}
                        }
                    }
                    else //No es ni Emisión ni Renovación:
                    {
                        //Con devolución al 100%: 
                        oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "false" });
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un error al validar reglas de Anulación para " + _cLinea, Ex);
            }
            return true;
        } 
        #endregion
    }
}
