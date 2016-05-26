using System;
using System.Collections.Generic;
using System.Windows.Forms;
using DocumentParser.models;
using DocumentParser.services;

namespace DocumentParser
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
//            TestingWordUtilsODK utilOdk = new TestingWordUtilsODK();
//            utilOdk.doSomething2();
//            ExperimentalServiceInterop experiment = new ExperimentalServiceInterop();
//            experiment.doSomething2();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new presentation.MainMenu());

        }
    }
}
