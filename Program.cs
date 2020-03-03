using Laserfiche.DocumentServices;
using Laserfiche.RepositoryAccess;
using System;
using System.Configuration;
using System.IO;
using System.Runtime.InteropServices;

/// <summary>
/// This Application takes a repository entry in Laserfiche that is a PDF electronic document,
/// downloads it to a temporary folder,
/// converts its pages to TIFF images,
/// and imports it into the repository under the same ID.
/// 
/// This effectively makes it a Laserfiche Page Generator because TIFF images are considered to be Laserfiche pages.
/// 
/// All other variables are stored in a separate App.config file.
/// 
/// INPUT PARAMETER : ID of entry that is a PDF (with no corresponding Laserfiche pages)
///          OUTPUT : # of pages.
/// </summary>
/// 

namespace Cni_Laserfiche_PageGenerator
{
    class Program
    {
        static int Main(string[] args)

        {
            if (args.Length == 0)
            {
                Logger.MessageToFile("Laserfiche PDF to TIFF did not receive an entry id.");
                return -1;
            }
            else
            {
                Logger.MessageToFile("Started looking at conversion....");
                /* CONSTANTS */
                string Laserfiche_server = ConfigurationManager.AppSettings["LaserficheServer"];
                string Laserfiche_repository = ConfigurationManager.AppSettings["LaserficheRepository"];
                string Laserfiche_user = ConfigurationManager.AppSettings["LaserficheUser"];
                string Laserfiche_password = ConfigurationManager.AppSettings["LaserfichePassword"];
                string temp_save_folder = ConfigurationManager.AppSettings["TemporarySaveFolder"];
                /* Ghostscript specific parameters */
                string tiff_image_resolution = "-r" + ConfigurationManager.AppSettings["TiffImageResolution"];
                string tiff_image_type = "-sDEVICE=" + ConfigurationManager.AppSettings["TiffImageType"];
                string inputMime;
                /* VARIABLES */
                int entry_id = Int32.Parse(args[0]);
                string message_id = "Laserfiche Entry#" + entry_id + "(" + Laserfiche_repository + ")";
                string message_id_eol = message_id + ".";

                /* LASERFICHE SESSION */
                RepositoryRegistration repositoryRegistration = new RepositoryRegistration(Laserfiche_server, Laserfiche_repository);
                Session LaserficheSession = new Session
                {
                    IsSecure = false
                };
                LaserficheSession.LogIn(Laserfiche_user, Laserfiche_password, repositoryRegistration);

                /* LASERFICHE OBJECTS */
                LaserficheReadStream electronicDocument = Document.ReadEdoc(entry_id, out inputMime, LaserficheSession);
                DocumentInfo document = Document.GetDocumentInfo(entry_id, LaserficheSession);
                EntryInfo entry = Entry.GetEntryInfo(entry_id, LaserficheSession);

                /* UNLOCK ENTRIES JUST IN CASE */
                document.Unlock();
                entry.Unlock();

                /* LENGTH OF PDF FILE / Byte Array */
                MemoryStream memoryStream = new MemoryStream();
                electronicDocument.CopyTo(memoryStream);
                byte[] byteArray = memoryStream.ToArray();

                /* LASERFICHE VARS - EXTRA INFORMATION */
                string fileExt = document.Extension.ToString();

                /* EXPORT PDF and convert to TIFFs */
                if (document.PageCount == 0 && inputMime == "application/pdf")
                {
                    Logger.MessageToFile("Started downloading PDF for " + message_id_eol);
                    /* Save PDF */
                    string pdf_file_path = temp_save_folder + document.Id.ToString() + '.' + fileExt;
                    FileStream fileStream = new FileStream(pdf_file_path, FileMode.Create, FileAccess.ReadWrite);
                    fileStream.Write(byteArray, 0, byteArray.Length);

                    fileStream.Close();

                    /* Ghostscript Variables */
                    var pdf_file_info = new FileInfo(pdf_file_path);
                    string tiff_file_path = Path.Combine(temp_save_folder, pdf_file_info.Name.Replace(pdf_file_info.Extension, ".tiff"));

                    /* delete the TIFF if it is already there....*/
                    if (File.Exists(tiff_file_path))
                    {
                        File.Delete(tiff_file_path);
                    }

                    /* Extract TIFF from PDF using Ghostscript DLL ... */
                    /* MORE INFORMATION ABOUT GHOSTSCRIPT : https://www.ghostscript.com/doc/current/Use.htm */
                    /* -dNOPAUSE -dBATCH keeps ghostscript from prompting user input */
                    /* -dSAFER enables access controls on files. https://www.ghostscript.com/doc/current/Use.htm#Safer */

                    try
                    {
                        Logger.MessageToFile("Started ghostscript conversion for " + message_id_eol);
                        string[] argv = new string[] { "PDF2TIFF", "-q", "-sOutputFile=" + tiff_file_path, "-dNOPAUSE", "-dBATCH", "-P-", "-dSAFER", tiff_image_type, tiff_image_resolution, pdf_file_path };
                        Ghostscript.Run(argv);
                    }
                    catch(Exception e)
                    {

                        Logger.MessageToFile("FAIL : Ghostscript conversion for " + message_id_eol);
                        Logger.MessageToFile(e.ToString());
                        Environment.Exit(0);
                    }
                    finally
                    {
                        Logger.MessageToFile("SUCCESS : Downloading / converting PDF to TIFF for " + message_id_eol);
                        try
                        {
                            DocumentImporter documentImporter = new DocumentImporter
                            {
                                Document = document
                            };
                            documentImporter.ImportImages(tiff_file_path);
                        }
                        catch (Exception e)
                        {
                            Logger.MessageToFile("FAIL : importing pages to " + message_id_eol);
                            Logger.MessageToFile(e.ToString());
                            Environment.Exit(0);
                        }
                        finally
                        {

                            Logger.MessageToFile("SUCCESS : imported pages to " + message_id_eol);
                            File.Delete(pdf_file_path);
                            File.Delete(tiff_file_path);
                            memoryStream.Close();
                            electronicDocument.Close();
                            document.Dispose();
                            entry.Dispose();
                            LaserficheSession.Close();
                            LaserficheSession.Discard();
                        }
                        
                    }
                    Logger.MessageToFile("SUCCESS : " + message_id + " uploaded with a total of " + document.PageCount + " pages.");
                    return document.PageCount;
                }
                else
                {
                    Logger.MessageToFile("BYPASS : " + message_id + " already had pages generated / " + document.PageCount + " pages.");
                    return document.PageCount;
                }
            }
        }
    }

    public class Ghostscript
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct GSVersion
        {
            public string product;
            public string copyright;
            public int revision;
            public int revisionDate;
        }

        [DllImport("gsdll32.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        private static extern int gsapi_revision(ref GSVersion version, int len);

        [DllImport("gsdll32.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        private static extern int gsapi_new_instance(ref IntPtr instance, IntPtr handle);

        [DllImport("gsdll32.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        private static extern int gsapi_init_with_args(IntPtr instance, int argc, string[] argv);

        [DllImport("gsdll32.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        private static extern int gsapi_exit(IntPtr instance);

        [DllImport("gsdll32.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        private static extern void gsapi_delete_instance(IntPtr instance);

        public static int GetVersion(ref GSVersion version)
        {
            return gsapi_revision(ref version, Marshal.SizeOf(version));
        }

        public static void Run(string[] argv)
        {
            var inst = IntPtr.Zero;
            int code = gsapi_new_instance(ref inst, IntPtr.Zero);
            if (code != 0)
            {
                return;
            }
            _ = Ghostscript.gsapi_init_with_args(inst, argv.Length, argv);
            gsapi_exit(inst);
            gsapi_delete_instance(inst);
        }

    }

}
