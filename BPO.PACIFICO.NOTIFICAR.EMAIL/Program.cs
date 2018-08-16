using everis.Ees.Proxy;
using everis.Ees.Proxy.Core;
using everis.Ees.Proxy.Services.Interfaces;
using Everis.Ees.Entities;
using Everis.Ees.Entities.Enums;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace BPO.Robot.Template.v3 //BPO.PACIFICO.NOTIFICAR.EMAIL
{
    /// <summary>
    /// Programa base para criação de Robôs para EES v3 2017 
    /// </summary>
    public class Program : IRobot
    {
        static BaseRobot<Program> _robot = null;
        public static int ReturnCode { get; private set; }
        private static ShowWindow ShowWindowOption
        {
            get
            {
                try
                {
                    return Convert.ToBoolean(Utils.GetKeyFromConfig("ShowWindow")) ? Everis.Ees.Entities.Enums.ShowWindow.Show : Everis.Ees.Entities.Enums.ShowWindow.Hide;
                }
                catch
                {
                    return Everis.Ees.Entities.Enums.ShowWindow.Hide;
                }
            }
        }

        #region Parametros
        private ParamRobotVirtualMachine actionOrderSuccessParam { get; set; }

        #endregion

        private static int Main(string[] args)
        {
            Utils.ShowWindow(ShowWindowOption);

            //Inicializando instancia do Robô
            _robot = new BaseRobot<Program>(args);

            //Inserindo parâmentros para mensageria
            _robot.InsertParamLogMessage(Utils.GetKeyFromConfig(Constants.KEY_LOG_MSG_NAME_STATE), Utils.GetKeyFromConfig(Constants.ROBOT_NAME));
            
            _robot.Start();

            return ReturnCode;
        }

        #region Start
        protected override void Start()
        {
            try
            {
                if (_robot.Tickets != null && _robot.Tickets.Count > 0)
                {
                    //Ticket Next State 
                    actionOrderSuccessParam = _robot.GetValueParamRobot("actionOrderSuccessParam");


                    foreach (var ticket in _robot.Tickets)
                    {
                        //TODO: fluxo principal do ticket



                        //Mudança de Fluxo do Ticket
                        SaveTicketNextState(ticket);
                    }
                }
                else
                {
                    LogEndStep(Constants.MSG_STEP_ENDED_KEY);
                }
            }
            catch (Exception ex)
            {
                LogFailProcess(Constants.MSG_ERROR_EVENT_PROCESS_KEY, ex);
                ReturnCode = ex.HResult;
            }

            LogEndStep(Constants.MSG_STEP_ENDED_KEY);
        }

        /// <summary>
        /// Método exemplo de mudança de estado para um ticket processado
        /// </summary>
        /// <param name="ticket"></param>
        private void SaveTicketNextState(Ticket ticket)
        {
            List<StateAction> actions = _robot.GetNextStateAction(ticket);
            if (actions != null && actions.Count > 0)
            {
                //por Action Order
                StateAction param = actions.FirstOrDefault(o =>
                                                    o.ActionOrder.ToString().Equals(actionOrderSuccessParam?.ValueParam));

                if (param != null)
                    _robot.SaveTicketNextState(ticket, param.Id);
            }
        }
        #endregion

        #region WindowHandler
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        const int SW_HIDE = 0;
        const int SW_SHOW = 5;
        #endregion
    }
}
