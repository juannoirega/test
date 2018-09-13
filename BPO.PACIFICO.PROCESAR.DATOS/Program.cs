﻿using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using Everis.Ees.Entities.Enums;
using Robot.Util.Nacar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RobotProcesarTicket
{
    class Program : IRobot
    {
        #region "PARÁMETROS"
        private static BaseRobot<Program> _oRobot = null;
        private static int _nDiasArrepentimiento;
        private static int _nDiasDesistimiento;
        private static string _cPolizaEmision;
        private static string _cPolizaRenovacion;
        private static string _cEstadoAnulacion;
        private static string _cLinea = string.Empty;
        private static string _cLineaAutos = string.Empty;
        private static string _cLineaLLPP = string.Empty;
        private static string _cLineaAlianzas = string.Empty;
        private static string _cLineaRRGG = string.Empty;
        private static string _cProceso = string.Empty;
        private static string _cTratamientoManual = string.Empty;
        private static string[] _procesos;
        private List<string> _productosAutos = new List<string>();
        private List<string> _productosRG = new List<string>();
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
                _oRobot = new BaseRobot<RobotProcesarTicket.Program>(args);
                _Funciones = new Functions();
                _oRobot.Start();
            }
            catch (Exception Ex) { Console.WriteLine(Ex.Message); }
        }
        protected void IniciarParametros()
        {
            _productosAutos.Add("Auto Modular");
            _productosAutos.Add("Auto a Medida");
            _productosAutos.Add("RCTPU (AX)");
            _productosRG.Add("MI01");
            _productosRG.Add("Hogar");
            _productosRG.Add("Accidentes Personales");
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

        protected override void Start()
        {
            if (_oRobot.Tickets.Count < 1)
                return;
            IniciarParametros();
            ObtenerParametros();
            LogStartStep(4);
            Inicio();
            foreach (Ticket oTicket in _oRobot.Tickets)
            {
                try
                {
                    _oMesaControl = _oRobot.GetNextStateAction(oTicket).First(a => a.ActionDescription == "Mesa de Control");
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
                _nDiasArrepentimiento = Convert.ToInt32(_oRobot.GetValueParamRobot("reglaDiasPolRenovadaAuto").ValueParam);
                _nDiasDesistimiento = Convert.ToInt32(_oRobot.GetValueParamRobot("reglaDiasPolNuevaAuto").ValueParam);
                _cPolizaEmision = "EMISION";
                _cPolizaRenovacion = "RENOVACION";
                _cEstadoAnulacion = "VIGENTE";
                _cLineaAutos = "AUTOS";
                _cLineaLLPP = "LLPP";
                _cLineaAlianzas = "ALIANZAS";
                _cLineaRRGG = "RRGG ";
                _procesos = _oRobot.GetValueParamRobot("reglaEstado").ValueParam.Split(',');
            }
            catch (Exception Ex) { LogFailStep(12, Ex); }
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

        //Metodo para generar los campos de  reglas
        private void InsertarValoresReglas(Ticket oTicketDatos, string msgConforme, string msgNoConforme, string msgObservacion)
        {
            //Reglas Conforme
            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.reglas_conforme, Value = msgConforme });

            //Reglas No Conforme
            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.reglas_no_conforme, Value = msgNoConforme });

            //Reglas Observación
            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.reglas_observacion, Value = msgObservacion });

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
                if (_productosAutos.Where(o => o == producto).FirstOrDefault() != null)
                {
                    _cLinea = "AUTOS";

                }
                else if (_productosRG.Where(o => o == producto).FirstOrDefault() != null)
                {
                    _cLinea = "RRGG";
                }
                else if (_productosAlianzas.Where(o => o == producto).FirstOrDefault() != null)
                {
                    _cLinea = "ALIANZAS";
                }
                if (_tProductosAlianzas.Where(o => o == tipoProducto).FirstOrDefault() != null)
                {
                    _cLinea = "ALIANZAS";

                }
                else if (_tProductosLPersonales.Where(o => o == tipoProducto).FirstOrDefault() != null)
                {
                    _cLinea = "LLPP";
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
                oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.linea, Value = _cLinea });

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
            int[] oCampos = new int[] {eesFields.Default.cuenta_nombre, eesFields.Default.asegurado_nombre, eesFields.Default.fecha_hora_de_email,
                                            eesFields.Default.tipo_poliza,eesFields.Default.poliza_fec_ini_vig,eesFields.Default.poliza_fec_fin_vig,
                                            eesFields.Default.poliza_est, eesFields.Default.poliza_tipo_vig};

            //Valida Línea de la Póliza:
            if (_cLinea == _cLineaAutos)
            {
                if (!ValidarDatosPoliza(oTicketDatos, oCampos))
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
            else if ((_cLinea == _cLineaAlianzas) || (_cLinea == _cLineaAlianzas) || (_cLinea == _cLineaAlianzas))
            {
                if (!ValidarDatosPoliza(oTicketDatos, oCampos))
                {
                    if (!ReglasAnulacionOtrasLineas(oTicketDatos))
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
                string estadoPoliza = oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_est).Value;
                Boolean _bFlagVigencia = (_procesos.Where(o => o == estadoPoliza).FirstOrDefault() == null);

                if (_bFlagVigencia) //Estado: VIGENTE. 
                {
                    TimeSpan nDiferencia = Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.fecha_hora_de_email).Value)
                                            - Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.poliza_fec_ini_vig).Value);

                    msgConforme = "La poliza se encuentra en estado: " + estadoPoliza + ". " + msgConforme;
                    //SE VERIFICA SINIESTRO
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.siniestros).Value != null)
                    {
                        msgConforme = "Se esta cumpliendo que la poliza no tiene siniestros. " + msgConforme;

                    }
                    else
                    {
                        msgNoConforme = "No se esta cumpliendo que la poliza no tiene siniestros. " + msgNoConforme;
                    }

                    //SE VERIFICA ENDOSOS
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.endosos).Value != null)
                    {
                        msgConforme = "Se esta cumpliendo que la poliza tiene endosos. " + msgConforme;
                    }
                    else
                    {
                        msgNoConforme = "No se esta cumpliendo que la poliza no tiene endosos. " + msgNoConforme;
                    }


                    //VERIFICA QUE SEA EMISIÓN:
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value.ToUpperInvariant() == _cPolizaEmision)
                    {
                        if (!(nDiferencia.Days > _nDiasDesistimiento))
                        {
                            msgConforme = "Se esta cumpliendo con la diferencia de dias de desestimiento. " + msgConforme;
                            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "1" });
                        }
                        else
                        {
                            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "0" });
                            msgNoConforme = "No se esta cumpliendo con la diferencia de dias de desestimiento. " + msgNoConforme;
                        }
                    }
                    else if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.tipo_poliza).Value.ToUpperInvariant() == _cPolizaRenovacion)//Es renovación:
                    {

                        if (!(nDiferencia.Days > _nDiasArrepentimiento))
                        {
                            msgConforme = "Se esta cumpliendo con la diferencia de dias de desestimiento." + msgConforme;
                            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "1" });
                        }
                        else
                        {
                            msgNoConforme = "No se esta cumpliendo con la diferencia de dias de desestimiento." + msgNoConforme;
                            oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "0" });
                        }
                    }
                    else //No es ni Emisión ni Renovación:
                    {
                        oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "0" });
                        InsertarValoresReglas(oTicketDatos, msgConforme, msgNoConforme, msgObservacion);
                        return false;
                    }
                }
                else
                {
                    msgNoConforme = " Se esta cumpliendo la regla por el hecho de que la poliza se encuentra en estado: " + estadoPoliza + msgNoConforme;
                    InsertarValoresReglas(oTicketDatos, msgConforme, msgNoConforme, msgObservacion);
                    return false;
                }
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un error al validar reglas de Anulación para " + _cLinea, Ex);
            }
            InsertarValoresReglas(oTicketDatos, msgConforme, msgNoConforme, msgObservacion);
            return true;
        }

        #endregion

        #region REGLAS DE VALIDACIÓN DE LALS OTRAS LÍNEAS
        //Reglas de validación para la línea Líneas Personales:
        private Boolean ReglasAnulacionOtrasLineas(Ticket oTicketDatos)
        {
            try
            {
                string estadoPoliza = oTicketDatos.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.poliza_est).Value;
                Boolean _bFlagVigencia = (_procesos.Where(o => o == estadoPoliza).FirstOrDefault() == null);

                if (_bFlagVigencia) //Estado: VIGENTE. 
                {
                    TimeSpan nDiferencia = Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.fecha_hora_de_email).Value)
                                            - Convert.ToDateTime(oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.poliza_fec_ini_vig).Value);

                    msgConforme = "La poliza se encuentra en estado: " + estadoPoliza + ". " + msgConforme;
                    //SE VERIFICA SINIESTRO
                    if (oTicketDatos.TicketValues.FirstOrDefault(o => o.FieldId == eesFields.Default.siniestros).Value != null)
                    {
                        msgConforme = "Se esta cumpliendo que la poliza no tiene siniestros. " + msgConforme;

                    }
                    else
                    {
                        msgNoConforme = "No se esta cumpliendo que la poliza no tiene siniestros. " + msgNoConforme;
                    }

                    if (!(nDiferencia.Days > _nDiasDesistimiento))
                    {
                        msgConforme = "Se esta cumpliendo con la diferencia de dias de desestimiento. " + msgConforme;
                        oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "1" });
                    }
                    else
                    {
                        oTicketDatos.TicketValues.Add(new TicketValue { ClonedValueOrder = null, TicketId = oTicketDatos.Id, FieldId = eesFields.Default.aplica_prorrata, Value = "0" });
                        msgNoConforme = "No se esta cumpliendo con la diferencia de dias de desestimiento. " + msgNoConforme;
                    }

                }
                else
                {
                    msgNoConforme = " Se esta cumpliendo la regla por el hecho de que la poliza se encuentra en estado: " + estadoPoliza + msgNoConforme;
                    InsertarValoresReglas(oTicketDatos, msgConforme, msgNoConforme, msgObservacion);
                    return false;
                }
            }
            catch (Exception Ex)
            {
                throw new Exception("Ocurrió un error al validar reglas de Anulación para " + _cLinea, Ex);
            }
            InsertarValoresReglas(oTicketDatos, msgConforme, msgNoConforme, msgObservacion);
            return true;

        }

        #endregion

    }
}
