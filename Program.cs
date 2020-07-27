using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace TrayGuard
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Pqm.CreateDocument();
            Application.Run(new frmLogin());
            //Application.Run(new frmModuleInTray());
        }
    }
}