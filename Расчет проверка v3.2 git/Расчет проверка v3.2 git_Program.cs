using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;
//using Microsoft.Office.Interop.Word;

namespace _2
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;//<<<-----
            //Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
           // Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-US");
            Application.Run(new Form1());
        }
    }
}
