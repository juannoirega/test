using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using Everis.Ees.Entities.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace BPO.PACIFICO.CREAR.TICKET.HIJO
{
    class Program : IRobot
    {
        #region Paremetros

        private static BaseRobot<Program> _robot = null;
        private static int _estadoError;
        private static int _estadoHijo;
        private static int _estadoPadre;
        private static int _fields;
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

            GetRobotParam();
            LogStartStep(4);
            foreach(Ticket ticket in _robot.Tickets)
                try
                {
                    ProcessaTicket(ticket);
                }
                catch(Exception ex)
                {
                    LogFailStep(42, ex);
                    EnviarPadre(ticket, true);
                }

            Environment.Exit(0);

        }

        private void GetRobotParam()
        {
          _estadoError = Convert.ToInt32(_robot.GetValueParamRobot("EstadoError").ValueParam);
          _estadoHijo = Convert.ToInt32(_robot.GetValueParamRobot("EstadoHijo").ValueParam);
          _estadoPadre = Convert.ToInt32(_robot.GetValueParamRobot("EstadoPadre").ValueParam);
          _fields = Convert.ToInt32(_robot.GetValueParamRobot("Fields").ValueParam);
        }

        private void ProcessaTicket(Ticket ticketPadre)
        {
            try
            {
                GeneraHijo(ticketPadre);
            }
            catch { throw new Exception("Error na gerecion de Hijos"); }

            EnviarPadre(ticketPadre, false);

        }
        private void EnviarPadre(Ticket ticketPadre, bool error)
        {
            try
            {
                var container = ODataContextWrapper.GetContainer();

                var ticket = container.Tickets.FirstOrDefault(o => o.Id == ticketPadre.Id);

                if (error)
                    ticket.StateId = Convert.ToInt32(ticketPadre.TicketValues.FirstOrDefault(o => o.FieldId == _estadoError).Value);
                else
                    ticket.StateId = Convert.ToInt32(ticketPadre.TicketValues.FirstOrDefault(o => o.FieldId == _estadoPadre).Value);

                container.UpdateObject(ticket);

                container.SaveChanges();
            }
            catch { throw new Exception("No fue posible realizar lo cambio de estado de el ticket Padre"); }
        }
        private void GeneraHijo(Ticket ticketPadre)
        {
            
            string [] campos = ticketPadre.TicketValues.FirstOrDefault(o => o.FieldId == _fields).Value.Split(Convert.ToChar(","));

            Ticket nuevoTicket = new Ticket { ParentId = ticketPadre.Id, Priority = PriorityType.Media, StateId = Convert.ToInt32(ticketPadre.TicketValues.FirstOrDefault(o => o.FieldId == _estadoHijo).Value) };


            _robot.SaveNewTicket(GeneraValuesHijo(ticketPadre, nuevoTicket, campos));
           
        }

        private Ticket GeneraValuesHijo(Ticket ticketPadre, Ticket nuevoTicket, string [] campos)
        {
            foreach (string campo in campos)
                if (ticketPadre.TicketValues.Where(o => o.FieldId == Convert.ToInt32(campo)).ToList().Count < 2)
                    nuevoTicket.TicketValues.Add(ticketPadre.TicketValues.FirstOrDefault(o => o.FieldId == Convert.ToInt32(campo)));
                else
                    foreach (TicketValue value in ticketPadre.TicketValues.Where(o => o.FieldId == Convert.ToInt32(campo)).ToList())
                        nuevoTicket.TicketValues.Add(value);

            return nuevoTicket;
        }
    }
}
