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

namespace BPO.PACIFICO.ProcesarDatos
{
    public class Program : IRobot
    {
        #region "PARÁMETROS"
        static BaseRobot<Program> _robot = null;
        private static int _nIdNomContratante = 19;
        private static int _nIdNomAsegurado = 18;
        private static int _nIdTipoPoliza = 26;
        private static int _nIdFechaInicioVigencia = 9;
        private static int _nIdFechaFinVigencia = 8;
        private static int _nIdProducto = 1034;
        private static int _nIdVistoBueno = 1026;
        private static int _nIdEstado = 1032;
        private static int _nIdTipoVigencia = 0;
        private static int _nIdAgente = 2;
        private static int _nIdFechaHoraEmail = 15;
        private static bool _bProrrata = false;
        //Parámetros del Robot Procesamiento de Datos:
        private static int _nDiasArrepentimiento = Convert.ToInt32(_robot.GetValueParamRobot("nDiasArrepentimiento").ValueParam);
        private static int _nDiasDesistimiento = Convert.ToInt32(_robot.GetValueParamRobot("nDiasDesistimiento").ValueParam);
        #endregion

        static void Main(string[] args)
        {
            _robot = new BaseRobot<Program>(args);
            _robot.Start();
        }

        protected override void Start()
        {
            if (_robot.Tickets.Count < 1)
            {
                return;
            }
            foreach (var ticket in _robot.Tickets)
            {
                ProcesarTicket(ticket);
            }

            //throw new NotImplementedException();    
        }

        //Inicia el procesamiento de datos:
        private void ProcesarTicket(Ticket ticket)
        {
            try
            {

                ValidaProducto(ticket);
            }
            catch (Exception)
            {

            }
        }

        private void ValidaProducto()
        {
            //Invoca a la validación correspondiente según producto:
            switch (Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdProducto).Value))
            {
                case 1:
                    //Valida las reglas generales para Autos:
                    if (Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdEstado).Value) == 1) //ID DEL ESTADO VIGENTE. 
                    {
                        //Valida las reglas específicas:
                        if (ValidarReglasAutos(ticket))
                        {

                        }
                    }
                    break;

                case 2:
                    //Valida las reglas generales para LLPP:
                    if (ValidarReglasLLPP())
                    {

                    }
                    break;

                case 3:
                    //Valida las reglas generales para BANA:
                    if (ValidarReglasAlianzas())
                    {

                    }
                    break;

                case 4:
                    //Valida las reglas generales para RRGG:
                    if (ValidarReglasRRGG())
                    {

                    }
                    break;
                default:
                    break;
            }
        }

        #region "REGLAS DE VALIDACIÓN"
        //1.- Autos:
        private bool ValidarReglasAutos(Ticket oTicket)
        {
            //VERIFICA QUE SEA EMISIÓN:
            if (Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdTipoPoliza).Value) == 2)
            {
                TimeSpan nDiferencia = Convert.ToDateTime(oTicket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdFechaHoraEmail).Value) - Convert.ToDateTime(oTicket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdFechaInicioVigencia).Value);
                int nDias = nDiferencia.Days;
                //Si ha superado los días de arrepentimiento:
                if (nDias > _nDiasArrepentimiento)
                {
                    if (oTicket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdVistoBueno).Value.Length <= 0)
                    {
                        return false;
                    }
                }
            }
            else
            {
                //No son emisiones:

            }



            try
            {
                ActualizarTicket();
                return true;
            }
            catch
            {
                Pausa();
            }
            return false;
        }

        //2.- Línas Generales:
        private bool ValidarReglasLLPP()
        {
            try
            {
                ActualizarTicket();
                return true;
            }
            catch
            {
                Pausa();
            }
            return false;
        }

        //3.- Banca y Alianzas:
        private bool ValidarReglasAlianzas()
        {
            try
            {
                ActualizarTicket();
                return true;
            }
            catch
            {
                Pausa();
            }
            return false;
        }

        //4.- Riesgos Generales:
        private bool ValidarReglasRRGG()
        {
            try
            {
                ActualizarTicket();
                return true;
            }
            catch
            {
                Pausa();
            }
            return false;
        }
        #endregion

        private void Pausa(double nTiempo = 1)
        {
            Thread.Sleep(1000 * Convert.ToInt32(nTiempo));
        }

        //Actualiza ticket con los nuevos datos para la anulación de póliza
        private void ActualizarTicket()
        {

        }



        //Concatena ubicación y nombre de imagen en archivo:
        private string EnlaceImagen(string cNombre)
        {
            return (String.Concat(@"Imagenes/", cNombre));
        }

    }
}
