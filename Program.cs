using System;
using System.Windows.Forms;

namespace CSharpFlexGrid
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Show splash screen first, then main menu
            Application.Run(new SplashScreen());
        }
    }
}