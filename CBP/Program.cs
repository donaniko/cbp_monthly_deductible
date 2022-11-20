using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CBP
{
    static class Program
    {
        public static string TACTICAL_SHEET_1 { get; set; }
        public static string TACTICAL_SHEET_2 { get; set; }
        public static string TITOLO_APP { get; set; }
        public static string TEMPLATE_ESTINZIONE_ANTICIPATA { get; set; }
        public static string TEMPLATE_RECESSO { get; set; }
        public static string TEMPLATE_DISDETTA { get; set; }
        /// <summary>
        /// Punto di ingresso principale dell'applicazione.
        /// </summary>
        [STAThread]
        static void Main()
        {            
            if (Environment.OSVersion.Version.Major >= 6)
            {
                SetProcessDPIAware();
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //var expire = new DateTime(2018, 12, 31);

            //if (DateTime.Now > expire)
            //{
            //    MessageBox.Show("Licenza scaduta. Contattare il fornitore del servizio", "Montly Generator");
            //    return;
            //}

            var banca = (ConfigurationManager.AppSettings["SoloBanca"] ?? "") == "true";
            TITOLO_APP = ConfigurationManager.AppSettings["Titolo"] ?? "";
            TACTICAL_SHEET_1 = ConfigurationManager.AppSettings["TACTICAL_SHEET_1"] ?? "";
            TACTICAL_SHEET_2 = ConfigurationManager.AppSettings["TACTICAL_SHEET_2"] ?? "";
            TEMPLATE_ESTINZIONE_ANTICIPATA = ConfigurationManager.AppSettings["TemplateEstinzioneAnticipata"] ?? "";
            TEMPLATE_RECESSO = ConfigurationManager.AppSettings["TemplateRecesso"] ?? "";
            TEMPLATE_DISDETTA = ConfigurationManager.AppSettings["TemplateDisdetta"] ?? "";

            if (banca)
            {
                Application.Run(new Bank());
            }
            else
            {
                Application.Run(new Main());
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();
    }
}
