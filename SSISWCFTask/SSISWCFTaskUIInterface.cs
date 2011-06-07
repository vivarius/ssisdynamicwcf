using System;
using System.Windows.Forms;
using Microsoft.SqlServer.Dts.Runtime;
using Microsoft.SqlServer.Dts.Runtime.Design;
using SSISWCFTask100;


namespace SSISWCFTask100
{
    class SSISWCFTaskUIInterface : IDtsTaskUI
    {
        #region Private Variables

        private TaskHost _taskHost;
        private Connections _connections;
        #endregion

        #region Constructor
        public SSISWCFTaskUIInterface()
        {
        }

        #endregion

        #region IDtsTaskUI Interface

        public void Initialize(TaskHost taskHost, IServiceProvider serviceProvider)
        {
            _taskHost = taskHost;
            IDtsConnectionService cs = serviceProvider.GetService(typeof(IDtsConnectionService)) as IDtsConnectionService;
            _connections = cs.GetConnections();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public ContainerControl GetView()
        {
            //Show the property window
            return new frmEditProperties(_taskHost, _connections);
        }

        public void Delete(IWin32Window parentWindow)
        {
        }

        public void New(IWin32Window parentWindow)
        {
        }

        #endregion
    }
}
