using GdPicture;
using System;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;


namespace TecnoDimOcr.OcrGdpicture
{
    public class Ocr
    {
        public Ocr()
        {
        }

        public static string castTopdf(string file, GdPictureImaging oGdPictureImaging, GdPicturePDF oGdPicturePDF, bool pdfa = true, string idioma = "por", string white = null, string titulo = null, string autor = null, string assunto = null, string palavrasChaves = null, string criador = null, int dpi = 250)
        {
            try
            {
                var exFile = Path.GetExtension(file).ToUpper();
                switch (exFile)
                {
                    case ".PDF":
                        #region PDF

                        String folderpdf = Guid.NewGuid().ToString();

                        string strpdf = "";

                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(file, false);


                        if (status == GdPictureStatus.OK)
                        {
                            int ident = 1;
                            int num1 = oGdPicturePDF.GetPageCount();
                            int num4 = 1;
                            string[] mergeArray = new string[num1];
                            Directory.CreateDirectory(folderpdf);
                            if (num1 > 0)
                            {

                                bool flagpdf = true;

                                while (num4 <= num1)
                                {

                                    oGdPicturePDF.SelectPage(num4);

                                    int numpdf1 = oGdPicturePDF.RenderPageToGdPictureImage(300, true);

                                    var docuemntoId = Guid.NewGuid();

                                    string sstr = string.Concat(Ocr.GetCurrentDirectory(), "\\GdPicture\\Idiomas");

                                    oGdPicturePDF.SaveToFile(folderpdf + @"\" + ident + "_" + docuemntoId + ".pdf");


                                    var id = oGdPictureImaging.PdfOCRStart(folderpdf + @"\" + ident + "_" + docuemntoId + ".pdf", true, "", "", "", "", "");

                                    oGdPictureImaging.PdfAddGdPictureImageToPdfOCR(id, numpdf1, "por", sstr, "");
                                    oGdPictureImaging.PdfOCRStop(id);


                                    mergeArray[num4 - 1] = folderpdf + @"\" + ident + "_" + docuemntoId + ".pdf";


                                    if (oGdPicturePDF.GetStat() == 0)
                                    {
                                        num4++;
                                        ident++;
                                    }
                                    else
                                    {
                                        flagpdf = false;
                                        break;
                                    }

                                    oGdPictureImaging.ReleaseGdPictureImage(numpdf1);


                                }
                                oGdPicturePDF.CloseDocument();


                                if (flagpdf)
                                {
                                    var strPdf1 = file.Replace(Path.GetExtension(file), ".pdf");
                                    oGdPicturePDF.MergeDocuments(mergeArray, strPdf1);
                                    strpdf = strPdf1;
                                    oGdPicturePDF.CloseDocument();
                                }

                                oGdPictureImaging.ClearGdPicture();

                                string str1pdf = ConfigurationManager.AppSettings["PastaDestinoTemp"].ToString();
                                if (File.Exists(str1pdf + @"\" + Path.GetFileName(strpdf)))
                                {
                                    File.Replace(strpdf, str1pdf + @"\" + Path.GetFileName(strpdf), null);
                                }
                                else
                                {
                                    File.Move(strpdf, str1pdf + @"\" + Path.GetFileName(strpdf));
                                }

                                var filefinal = str1pdf + @"\" + Path.GetFileName(strpdf);
                                foreach (var item in Directory.GetFiles(folderpdf))
                                {
                                    File.Delete(item);
                                }
                                Directory.Delete(folderpdf);
                                file = filefinal;
                            }
                            else
                            {
                                oGdPicturePDF.SelectPage(num4);
                                int numpdf = oGdPicturePDF.RenderPageToGdPictureImage(300, true);
                                var docuemntoId = Guid.NewGuid();
                                string sstr = string.Concat(Ocr.GetCurrentDirectory(), "\\GdPicture\\Idiomas");
                                oGdPictureImaging.SaveAsPDFOCR(numpdf, folderpdf + @"\" + docuemntoId + ".pdf", idioma, sstr, white, pdfa, titulo, autor, assunto, palavrasChaves, criador);


                                var strPdf = file.Replace(Path.GetExtension(file), ".pdf");
                                oGdPictureImaging.ReleaseGdPictureImage(numpdf);




                                oGdPicturePDF.MergeDocuments(System.IO.Directory.GetFiles(folderpdf), strPdf);
                                strpdf = strPdf;

                                oGdPictureImaging.ClearGdPicture();


                                string str1tif = ConfigurationManager.AppSettings["PastaDestinoTemp"].ToString();
                                if (File.Exists(str1tif + @"\" + Path.GetFileName(strpdf)))
                                {
                                    File.Replace(strpdf, str1tif + @"\" + Path.GetFileName(strpdf), null);
                                }
                                else
                                {
                                    File.Move(strpdf, str1tif + @"\" + Path.GetFileName(strpdf));
                                }

                                var filefinal = str1tif + @"\" + Path.GetFileName(strpdf);
                                foreach (var item in Directory.GetFiles(folderpdf))
                                {
                                    File.Delete(item);
                                }
                                Directory.Delete(folderpdf);
                                file = filefinal;
                            }

                        }

                        #endregion
                        break;

                    case (".TIF"):
                        String folder = Guid.NewGuid().ToString();

                        string str = "";
                        oGdPictureImaging.TiffOpenMultiPageForWrite(false);
                        int num = oGdPictureImaging.CreateGdPictureImageFromFile(file);
                        if (num != 0)
                        {
                            int ident = 1;
                            Directory.CreateDirectory(folder);

                            if (oGdPictureImaging.TiffIsMultiPage(num))
                            {
                                int num1 = oGdPictureImaging.TiffGetPageCount(num);
                                bool flag = true;
                                int num4 = 1;
                                string[] mergeArray = new string[num1];
                                while (num4 <= num1)
                                {

                                    oGdPictureImaging.TiffSelectPage(num, num4);
                                    // oGdPicturePDF.AddImageFromGdPictureImage(num, false, true);
                                    oGdPicturePDF.NewPDF();
                                    var docuemntoId = Guid.NewGuid();
                                    oGdPicturePDF.SaveToFile(folder + @"\" + ident + "_" + docuemntoId + ".pdf");
                                    var id = oGdPictureImaging.PdfOCRStart(folder + @"\" + ident + "_" + docuemntoId + ".pdf", true, "", "", "", "", "");

                                    string sstr = string.Concat(Ocr.GetCurrentDirectory(), "\\GdPicture\\Idiomas");
                                    oGdPictureImaging.PdfAddGdPictureImageToPdfOCR(id, num, "por", sstr, "");
                                    oGdPictureImaging.PdfOCRStop(id);
                                    //     oGdPictureImaging.SaveAsPDFOCR(num4, @"C:\Processamento" + @"\" + docuemntoId + ".pdf", idioma, sstr, "", true, titulo, autor, assunto, palavrasChaves, criador);

                                    oGdPicturePDF.CloseDocument();
                                    mergeArray[num4 - 1] = folder + @"\" + ident + "_" + docuemntoId + ".pdf";
                                    if (oGdPicturePDF.GetStat() == 0)
                                    {
                                        num4++;
                                        ident++;
                                    }
                                    else
                                    {
                                        flag = false;
                                        break;
                                    }
                                    //      oGdPictureImaging.ReleaseGdPictureImage(num);

                                }


                                if (flag)
                                {
                                    var strPdf = file.Replace(Path.GetExtension(file), ".pdf");

                                    oGdPicturePDF.MergeDocuments(mergeArray, strPdf);
                                    str = strPdf;
                                }

                                oGdPictureImaging.ReleaseGdPictureImage(num);
                                oGdPictureImaging.ClearGdPicture();

                                string str1tif = ConfigurationManager.AppSettings["PastaDestinoTemp"].ToString();
                                if (File.Exists(str1tif + @"\" + Path.GetFileName(str)))
                                {
                                    File.Replace(str, str1tif + @"\" + Path.GetFileName(str), null);
                                    File.Delete(file);
                                }
                                else
                                {
                                    File.Move(str, str1tif + @"\" + Path.GetFileName(str));
                                    File.Delete(file);
                                }

                                var filefinal = str1tif + @"\" + Path.GetFileName(str);
                                foreach (var item in Directory.GetFiles(folder))
                                {
                                    File.Delete(item);
                                }
                                Directory.Delete(folder);
                                file = filefinal;
                            }
                            else
                            {
                                var docuemntoId = Guid.NewGuid();
                                string sstr = string.Concat(Ocr.GetCurrentDirectory(), "\\GdPicture\\Idiomas");

                                oGdPicturePDF.NewPDF();
                                oGdPicturePDF.SaveToFile(folder + @"\" + ident + "_" + docuemntoId + ".pdf");
                                var id = oGdPictureImaging.PdfOCRStart(folder + @"\" + ident + "_" + docuemntoId + ".pdf", true, "", "", "", "", "");
                                oGdPictureImaging.PdfAddGdPictureImageToPdfOCR(id, num, "por", sstr, "");
                                oGdPictureImaging.PdfOCRStop(id);
                                oGdPicturePDF.CloseDocument();
                                
                                //oGdPictureImaging.SaveAsPDFOCR(num, folder + @"\" + docuemntoId + ".pdf", idioma, sstr, white, pdfa, titulo, autor, assunto, palavrasChaves, criador);


                                var strPdf = file.Replace(Path.GetExtension(file), ".pdf");

                                oGdPicturePDF.MergeDocuments(System.IO.Directory.GetFiles(folder), strPdf);

                                str = strPdf;

                                oGdPictureImaging.ReleaseGdPictureImage(num);
                                oGdPicturePDF.CloseDocument();

                                oGdPictureImaging.ClearGdPicture();
                                string str1tif = ConfigurationManager.AppSettings["PastaDestinoTemp"].ToString();
                                if (File.Exists(str1tif + @"\" + Path.GetFileName(str)))
                                {
                                    File.Replace(str, str1tif + @"\" + Path.GetFileName(str), null);
                                    File.Delete(file);
                                }
                                else
                                {
                                    File.Move(str, str1tif + @"\" + Path.GetFileName(str));
                                    File.Delete(file);
                                }

                                var filefinal = str1tif + @"\" + Path.GetFileName(str);
                                foreach (var item in Directory.GetFiles(folder))
                                {
                                    File.Delete(item);
                                }
                                Directory.Delete(folder);
                                file = filefinal;
                            }
                        }

                        break;

                    case ".TIFF":
                        String folder2 = Guid.NewGuid().ToString();

                        string str2 = "";
                        oGdPictureImaging.TiffOpenMultiPageForWrite(false);
                        int num2 = oGdPictureImaging.CreateGdPictureImageFromFile(file);
                        if (num2 != 0)
                        {
                            int ident = 1;

                            Directory.CreateDirectory(folder2);
                            if (oGdPictureImaging.TiffIsMultiPage(num2))
                            {
                                int num1 = oGdPictureImaging.TiffGetPageCount(num2);
                                bool flag = true;
                                int num3 = 1;
                                string[] mergeArray = new string[num1];
                                while (num3 <= num1)
                                {

                                    oGdPictureImaging.TiffSelectPage(num2, num3);
                                    oGdPicturePDF.NewPDF();
                                    oGdPicturePDF.AddImageFromGdPictureImage(num2, false, true);

                                    var docuemntoId = Guid.NewGuid();
                                    // oGdPicturePDF.SaveToFile(folder + @"\" + docuemntoId + ".pdf");
                                    oGdPicturePDF.SaveToFile(folder2 + @"\" + ident + "_" + docuemntoId + ".pdf");
                                    //   var id = oGdPictureImaging.PdfOCRStart(folder + @"\" + docuemntoId + ".pdf", true, "", "", "", "", "");
                                    //oGdPictureImaging.PdfAddGdPictureImageToPdfOCR(id, num, "por", str, "");
                                    string sstr = string.Concat(Ocr.GetCurrentDirectory(), "\\GdPicture\\Idiomas");

                                    var id = oGdPictureImaging.PdfOCRStart(folder2 + @"\" + ident + "_" + docuemntoId + ".pdf", true, "", "", "", "", "");
                                    oGdPictureImaging.PdfAddGdPictureImageToPdfOCR(id, num2, "por", sstr, "");
                                    oGdPictureImaging.PdfOCRStop(id);
                                    oGdPicturePDF.CloseDocument();
                                    //oGdPictureImaging.SaveAsPDFOCR(num3, folder2 + @"\" + docuemntoId + ".pdf", idioma, sstr, white, pdfa, titulo, autor, assunto, palavrasChaves, criador);
                                    mergeArray[num3 - 1] = folder2 + @"\" + ident + "_" + docuemntoId + ".pdf";


                                    if (oGdPicturePDF.GetStat() == 0)
                                    {
                                        num3++;
                                        ident++;
                                    }
                                    else
                                    {
                                        flag = false;
                                        break;
                                    }
                                    //oGdPictureImaging.ReleaseGdPictureImage(num2);

                                }

                                if (flag)
                                {
                                    var strPdf = file.Replace(Path.GetExtension(file), ".pdf");

                                    oGdPicturePDF.MergeDocuments(mergeArray, strPdf);
                                    str2 = strPdf;
                                }
                                oGdPictureImaging.ReleaseGdPictureImage(num2);


                                oGdPictureImaging.ClearGdPicture();
                                string str1tiff = ConfigurationManager.AppSettings["PastaDestinoTemp"].ToString();
                                if (File.Exists(str1tiff + @"\" + Path.GetFileName(str2)))
                                {
                                    File.Replace(str2, str1tiff + @"\" + Path.GetFileName(str2), null);
                                    File.Delete(file);
                                }
                                else
                                {
                                    File.Move(str2, str1tiff + @"\" + Path.GetFileName(str2));
                                    File.Delete(file);

                                }

                                var filefinal2 = str1tiff + @"\" + Path.GetFileName(str2);
                                foreach (var item in Directory.GetFiles(folder2))
                                {
                                    File.Delete(item);
                                }
                                Directory.Delete(folder2);
                                file = filefinal2;
                            }
                            else
                            {
                                var docuemntoId = Guid.NewGuid();
                                string sstr = string.Concat(Ocr.GetCurrentDirectory(), "\\GdPicture\\Idiomas");

                                oGdPicturePDF.NewPDF();
                                var id = oGdPictureImaging.PdfOCRStart(folder2 + @"\" + ident + "_" + docuemntoId + ".pdf", true, "", "", "", "", "");
                                oGdPicturePDF.SaveToFile(folder2 + @"\" + ident + "_" + docuemntoId + ".pdf");
                                oGdPictureImaging.PdfAddGdPictureImageToPdfOCR(id, num2, "por", sstr, "");
                                oGdPictureImaging.PdfOCRStop(id);
                                oGdPicturePDF.CloseDocument();

                                var strPdf = file.Replace(Path.GetExtension(file), ".pdf");

                                oGdPicturePDF.MergeDocuments(System.IO.Directory.GetFiles(folder2), strPdf);

                                str2 = strPdf;

                                oGdPictureImaging.ReleaseGdPictureImage(num2);
                                oGdPicturePDF.CloseDocument();

                                oGdPictureImaging.ClearGdPicture();
                                string str1tiff = ConfigurationManager.AppSettings["PastaDestinoTemp"].ToString();
                                if (File.Exists(str1tiff + @"\" + Path.GetFileName(str2)))
                                {
                                    File.Replace(str2, str1tiff + @"\" + Path.GetFileName(str2), null);
                                    File.Delete(file);
                                }
                                else
                                {
                                    File.Move(str2, str1tiff + @"\" + Path.GetFileName(str2));
                                    File.Delete(file);
                                }
                                var filefinal2 = str1tiff + @"\" + Path.GetFileName(str2);
                                foreach (var item in Directory.GetFiles(folder2))
                                {
                                    File.Delete(item);
                                }
                                Directory.Delete(folder2);
                                file = filefinal2;
                            }
                        }

                        break;

                }
            }
            catch (Exception ex)
            {
                var logpath = ConfigurationManager.AppSettings["PastaDestinoLog"].ToString();
                File.AppendAllText(logpath + @"\" + "log.txt", ex.ToString());

                //  Console.WriteLine(ex.Message);

            }

            return file;
        }

        private static string GetCurrentDirectory()
        {
            string absolutePath = (new Uri(Assembly.GetExecutingAssembly().CodeBase)).AbsolutePath;
            string fullName = (new DirectoryInfo(Path.GetDirectoryName(absolutePath))).FullName;
            return Uri.UnescapeDataString(fullName);

        }

        public static bool IsFileLocked(string filePath)
        {
            bool flag;
            try
            {
                using (FileStream fileStream = File.Open(filePath, FileMode.Open))
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
    }
}