using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using TecnoDimOcr.OcrGdpicture;

namespace TecnoDimOcr
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {

            try
            {
                //new ServiceOcrTecnodim().teste();// utar(null, null);
                ServiceBase.Run(new ServiceBase[] { new ServiceOcrTecnodim() });

            }
            catch (Exception ex)
            {
                var logpath = ConfigurationManager.AppSettings["PastaDestinoLog"].ToString();
                File.AppendAllText(logpath + @"\" + "log.txt", ex.ToString());
                
                //  Console.WriteLine(ex.Message);
                
            }


        }
    }
}
