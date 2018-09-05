using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using Robot.Util.Nacar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPO.PACIFICO.PROCESARDATOS.APP
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
                _oRobot = new BaseRobot<BPO.PACIFICO.PROCESARDATOS.APP.Program>(args);
                _Funciones = new Functions();
                _oRobot.Start();
            }
            catch (Exception Ex) { Console.WriteLine(Ex.Message); }
        }

        protected override void Start()
        {
            if (_oRobot.Tickets.Count < 1)
                return;

            //ObtenerParametros();
            LogStartStep(4);
            //Inicio();
            foreach (Ticket oTicket in _oRobot.Tickets)
            {
                try
                {
                    _cProceso = _Funciones.ObtenerValorDominio(oTicket, Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(a => a.FieldId == eesFields.Default.tipo_proceso).Value));
                    _oMesaControl = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdMesaControl);
                    _oPantallaValidacion = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdPantallaValidacion);
                    _oNotificacion = _oRobot.GetNextStateAction(oTicket).First(a => a.DestinationStateId == _nIdNotificacion);
                    //ProcesosEndoso(oTicket);
                    //ProcesarTicket(oTicket);
                }
                catch (Exception Ex)
                {
                    //CambiarEstadoTicket(oTicket, _oMesaControl, Ex.Message);
                    LogFailStep(12, Ex);
                    return;
                }
            }
        }

    }
}
