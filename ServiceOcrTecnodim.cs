

using GdPicture;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using TecnoDimOcr.OcrGdpicture;


namespace TecnoDimOcr
{
    partial class ServiceOcrTecnodim : ServiceBase
    {

        public ServiceOcrTecnodim()
        {
            oGdPicturePDF.SetLicenseNumber("4118106456693265856441854");
            oGdPictureImaging.SetLicenseNumber("4118106456693265856441854");
            InitializeComponent();
            this.ServiceName = "ServiceOcrTecnodim";

        }
        private GdPictureImaging oGdPictureImaging = new GdPictureImaging();

        private GdPicturePDF oGdPicturePDF = new GdPicturePDF();


        public string executarOcr(string arquivo)
        {
            try
            {

                string documento = Ocr.castTopdf(arquivo, oGdPictureImaging, oGdPicturePDF);
                getPastaDestino(documento, arquivo);
            }
            catch (Exception ex)
            {
                var logpath = ConfigurationManager.AppSettings["PastaDestinoLog"].ToString();
                File.AppendAllText(logpath + @"\" + "log.txt", ex.ToString());

                //  Console.WriteLine(ex.Message);

            }
            return "";
        }

        public void getPastaDestino(string arquivo, string file)
        {
            try
            {
                if (Path.GetExtension(file).ToUpper() == ".PDF")
                {
                    var pdfname = Path.GetFileName(file).ToString();
                    string str1PDF = ConfigurationManager.AppSettings["PastaDestinoTemp"].ToString();
                    arquivo = str1PDF + @"\" + pdfname;

                    string strpdf = "";
                    string str1pdf = Path.GetFileName(arquivo).ToString();


                    string[] strArrayspdf = Path.GetFileName(arquivo).ToString().Split(new char[] { '_' });
                    if (strArrayspdf.Length != 0)
                    {
                        strpdf = strArrayspdf[0];
                    }
                    string str2pdf = ConfigurationManager.AppSettings["PastaDestino"].ToString();
                    str2pdf = string.Concat(str2pdf, "\\", strpdf);
                    if (!Directory.Exists(str2pdf))
                    {
                        str2pdf = ConfigurationManager.AppSettings["PastaDestinoNAOMAPEADO"].ToString();
                    }
                    if (File.Exists(str2pdf + @"\" + str1pdf))
                    {
                        File.Replace(arquivo, str2pdf + @"\" + str1pdf, null);
                    }
                    else
                    {
                        File.Move(arquivo, str2pdf + @"\" + str1pdf);
                    }
                    File.Delete(file);
                }
                else
                {
                    string str = "";
                    string str1 = Path.GetFileName(arquivo).ToString();


                    string[] strArrays = Path.GetFileName(arquivo).ToString().Split(new char[] { '_' });
                    if (strArrays.Length != 0)
                    {
                        str = strArrays[0];
                    }
                    string str2 = ConfigurationManager.AppSettings["PastaDestino"].ToString();
                    str2 = string.Concat(str2, "\\", str);
                    if (!Directory.Exists(str2))
                    {
                        str2 = ConfigurationManager.AppSettings["PastaDestinoNAOMAPEADO"].ToString();
                    }
                    if (File.Exists(str2 + @"\" + str1))
                    {
                        File.Replace(arquivo, str2 + @"\" + str1, null);
                    }
                    else
                    {
                        File.Move(arquivo, string.Concat(str2, "\\", str1));
                    }
                    File.Delete(file);
                }

            }
            catch (Exception ex)
            {
                var logpath = ConfigurationManager.AppSettings["PastaDestinoLog"].ToString();
                File.AppendAllText(logpath + @"\" + "log.txt", ex.ToString());

                //  Console.WriteLine(ex.Message);

            }
        }

        public bool IsFileLocked(string filename)
        {
            bool flag;
            try
            {
                using (FileStream fileStream = File.Open(filename, FileMode.Open))
                {
                }
            }
            catch (IOException oException)
            {
                var logpath = ConfigurationManager.AppSettings["PastaDestinoLog"].ToString();
                File.AppendAllText(logpath + @"\" + "log.txt", oException.ToString());

                int hRForException = Marshal.GetHRForException(oException) & 65535;
                flag = (hRForException == 32 ? true : hRForException == 33);
                return flag;
            }
            flag = false;
            return flag;
        }

        private static string GetCurrentDirectory()
        {
            string absolutePath = (new Uri(Assembly.GetExecutingAssembly().CodeBase)).AbsolutePath;
            string fullName = (new DirectoryInfo(Path.GetDirectoryName(absolutePath))).FullName;
            return Uri.UnescapeDataString(fullName);
        }


        public void executarEvento(object sender, FileSystemEventArgs e)
        {
            

            


           while(Ocr.IsFileLocked(e.FullPath))
            {

            }
            File.AppendAllText("logprocessado.txt", e.FullPath);
            executar(e.FullPath);
        }



        public void executar(string CAMINHO)
        {
            try
            {
                criarpastas();
                string str1 = ConfigurationManager.AppSettings["pastaEntradaTemp"].ToString();

                var item = CAMINHO;
                if (Path.GetExtension(item).ToUpper() == ".TIFF" || Path.GetExtension(item).ToUpper() == ".TIF" || Path.GetExtension(item).ToUpper() == ".PDF")
                {
                    if (!Ocr.IsFileLocked(item))
                    {

                        if (File.Exists(str1 + @"" + Path.GetFileName(item)))
                        {

                            File.Delete(str1 + @"" + Path.GetFileName(item));
                        }

                        File.Move(item, string.Concat(str1, "\\", Path.GetFileName(item)));
                    }
                }



                item = string.Concat(str1, "\\", Path.GetFileName(item));

                if (Path.GetExtension(item).ToUpper() == ".TIFF" || Path.GetExtension(item).ToUpper() == ".TIF" || Path.GetExtension(item).ToUpper() == ".PDF")
                {

                    try
                    {
                        executarOcr(item);
                    }
                    catch (Exception ex)
                    {
                        var logpath = ConfigurationManager.AppSettings["PastaDestinoLog"].ToString();
                        File.AppendAllText(logpath + @"\" + "log.txt", ex.ToString());

                    }

                }
            }


            catch (Exception ex)
            {
                var logpath = ConfigurationManager.AppSettings["PastaDestinoLog"].ToString();
                File.AppendAllText(logpath + @"\" + "log.txt", ex.ToString());

                //  Console.WriteLine(ex.Message);

            }


        }

        private void criarpastas()
        {
            var pastaEntrada = ConfigurationManager.AppSettings["pastaEntrada"].ToString();
            if (!Directory.Exists(pastaEntrada))
            {
                Directory.CreateDirectory(pastaEntrada);
            }
            var pastaEntradaTemp = ConfigurationManager.AppSettings["pastaEntradaTemp"].ToString();
            if (!Directory.Exists(pastaEntradaTemp))
            {
                Directory.CreateDirectory(pastaEntradaTemp);
            }
            var PastaDestinoRaiz = ConfigurationManager.AppSettings["PastaDestinoRaiz"].ToString();
            if (!Directory.Exists(PastaDestinoRaiz))
            {
                Directory.CreateDirectory(PastaDestinoRaiz);
            }
            var PastaDestinoTemp = ConfigurationManager.AppSettings["PastaDestinoTemp"].ToString();
            if (!Directory.Exists(PastaDestinoTemp))
            {
                Directory.CreateDirectory(PastaDestinoTemp);
            }
            var PastaDestinoNAOMAPEADO = ConfigurationManager.AppSettings["PastaDestinoNAOMAPEADO"].ToString();
            if (!Directory.Exists(PastaDestinoNAOMAPEADO))
            {
                Directory.CreateDirectory(PastaDestinoNAOMAPEADO);
            }
            var PastaDestinoLog = ConfigurationManager.AppSettings["PastaDestinoLog"].ToString();
            if (!Directory.Exists(PastaDestinoLog))
            {
                Directory.CreateDirectory(PastaDestinoLog);
            }
        }

        public void teste()
        {

            criarpastas();
            string path = ConfigurationManager.AppSettings["pastaEntrada"].ToString();



            foreach (var item in Directory.GetFiles(path))
            {
                executar(item);
                File.Delete(item);
            }


            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = path;
            watcher.NotifyFilter = NotifyFilters.LastWrite;
            watcher.Filter = "*.*";
            watcher.Created += new FileSystemEventHandler(executarEvento);
            watcher.EnableRaisingEvents = true;

          


        }


        protected override void OnStart(string[] args)
        {
            try
            {



                
                string path = ConfigurationManager.AppSettings["pastaEntrada"].ToString();



                foreach (var item in Directory.GetFiles(path))
                {
                    executar(item);
                    File.Delete(item);
                }


                FileSystemWatcher watcher = new FileSystemWatcher();
                watcher.Path = path;
                watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite
                   | NotifyFilters.FileName | NotifyFilters.DirectoryName;
                watcher.Filter = "*.*";
                watcher.Created += new FileSystemEventHandler(executarEvento);
                watcher.EnableRaisingEvents = true;



            }
            catch (Exception ex)
            {
                var logpath = ConfigurationManager.AppSettings["PastaDestinoLog"].ToString();
                File.AppendAllText(logpath + @"\" + "log.txt", ex.ToString());

                //  Console.WriteLine(ex.Message);

            }
        }

        protected override void OnStop()
        {

        }
    }
}
