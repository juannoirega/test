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
        private static int _nIdTipoLinea = 1;
        private static int _nIdSinCastigo = 2;
        private static int _nIdSinSiniestro = 3;
        private static int _nIdEstado = 4;
        private static int _nIdTipoAnulacion = 5;
        private static int _nIdNomContratante = 6;
        private static int _nIdNomContratanteCorreo = 7;
        private static int _nIdNomAsegurado = 8;
        private static int _nIdNomAseguradoCorreo = 9;
        private static int _nIdVistoBueno = 10;
        private static int _nIdFechaAnulacion = 11;
        private static int _nIdFechaInicio = 12;
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

        private void ProcesarTicket(Ticket ticket)
        {
            //Invoca a la validación correspondiente:
            switch (Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdTipoLinea).Value))
            {
                case 1:
                    //Valida las reglas generales para Autos:
                    if (Convert.ToBoolean(ticket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdSinCastigo).Value) &&
                        Convert.ToBoolean(ticket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdSinSiniestro).Value) &&
                        Convert.ToInt32(ticket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdEstado).Value) == 1) //ID DEL ESTADO VIGENTE. 
                    {
                        //Valida las reglas específicas:
                        if (ValidarReglasAutos(ticket))
                        {

                        }
                    }
                    break;

                case 2:
                    if (ValidarReglasLLPP())
                    {

                    }
                    break;

                case 3:
                    if (ValidarReglasAlianzas())
                    {

                    }
                    break;

                case 4:
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
            switch (Convert.ToInt32(oTicket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdTipoAnulacion).Value))
            {
                case 1:
                    if (oTicket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdNomContratante).Value == oTicket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdNomContratanteCorreo).Value ||
                        oTicket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdNomAsegurado).Value == oTicket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdNomAseguradoCorreo).Value)
                    {
                        return true;
                    }
                    break;
                case 2:
                    if (Convert.ToBoolean(oTicket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdVistoBueno).Value))
                    {
                        return true;
                    }
                    break;
                case 3:
                    if (Convert.ToInt32(Convert.ToDateTime(oTicket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdFechaAnulacion).Value) - Convert.ToDateTime(oTicket.TicketValues.FirstOrDefault(o => o.FieldId == _nIdFechaInicio).Value))
                        > Convert.ToInt32(_robot.GetValueParamRobot("nDiasProrrata").ValueParam))
                    {

                    }
                    break;
                default:
                    break;
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
