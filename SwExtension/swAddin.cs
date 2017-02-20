using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace SwExtension
{

    [Guid("F1F96231-4928-4CAC-A81E-CB5B2AB4B3BE"), ComVisible(true)]
    [SwAddin(Description = "Extension des fonctions Sw", Title = "Sw Extension", LoadAtStartup = true)]
    public partial class swAddin : ISwAddin
    {
        private static SldWorks _SwApp;

        private int _AddInCookie;

        private TaskpaneView _TaskpaneOngletLog;
        private OngletLog _OngletLog;

        private TaskpaneView _TaskpaneOngletParametres;
        private OngletParametres _OngletParametres;

        public static SldWorks SwApp { get { return _SwApp; } }

        bool ISwAddin.ConnectToSW(object ThisSW, int Cookie)
        {
            try
            {
                Log.Demarrer();

                _AddInCookie = Cookie;

                // On initialise les outils d'extension SW
                App.Sw = (SldWorks)ThisSW;

                _SwApp = (SldWorks)ThisSW;
                _SwApp.SetAddinCallbackInfo2(0, this, _AddInCookie);
                CreerCmdMgr();
                CreerTaskpane();
                CreerMenusEtOnglets();
                AddEventHooks();
                return true;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }

            return false;
        }

        bool ISwAddin.DisconnectFromSW()
        {
            SupprimerTaskpane();
            SupprimerMenus();
            RemoveEventHooks();
            SupprimerPMP();

            // On nettoie les références à SW

            _SwApp = null;
            App.Sw = null;
            Log.Stopper();
            GC.Collect();
            return true;
        }

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
            Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

            String keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
            Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
            addinkey.SetValue(null, 0);
            addinkey.SetValue("Description", "Extension des fonctions Sw");
            addinkey.SetValue("Title", "Sw Extension");
            keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
            addinkey = hkcu.CreateSubKey(keyname);
            addinkey.SetValue(null, 1);
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;
                String keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);
                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch { }
        }

        public void AddEventHooks()
        {
        }

        public void RemoveEventHooks()
        {
        }

        public void CreerTaskpane()
        {
            String codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            String CheminDossier = Path.GetDirectoryName(Uri.UnescapeDataString(uri.Path));

            String CheminImageOngletLog = Path.Combine(CheminDossier, "Icon_OngletLog." + ImageFormat.Bmp.ToString().ToLower());
            String CheminImageOngletParametre = Path.Combine(CheminDossier, "Icon_OngletParametre." + ImageFormat.Bmp.ToString().ToLower());

            Image ImgOngletLog = "L".eConvertirEnBmp(18, 16);
            Image ImgOngletParametre = "P".eConvertirEnBmp(18, 16);

            ImgOngletLog.Save(CheminImageOngletLog, ImageFormat.Bmp);
            ImgOngletParametre.Save(CheminImageOngletParametre, ImageFormat.Bmp);

            _TaskpaneOngletParametres = _SwApp.CreateTaskpaneView2(CheminImageOngletParametre, "Parametres");
            _TaskpaneOngletLog = _SwApp.CreateTaskpaneView2(CheminImageOngletLog, "Log");

            _OngletLog = new OngletLog();
            _OngletParametres = new OngletParametres(SwApp);

            _TaskpaneOngletLog.DisplayWindowFromHandlex64(_OngletLog.Handle.ToInt64());
            _TaskpaneOngletParametres.DisplayWindowFromHandlex64(_OngletParametres.Handle.ToInt64());

            _SwApp.ActiveDocChangeNotify += _OngletParametres.ActiveDocChange;
            _SwApp.FileCloseNotify += delegate (String nomFichier, int raison) { return _OngletParametres.ActiveDocChange(); };

            _SwApp.ActiveDocChangeNotify += _OngletParametres.Rechercher_Propriete_Modele;
            _SwApp.FileCloseNotify += delegate (String nomFichier, int raison) { return _OngletParametres.Rechercher_Propriete_Modele(); };

            WindowLog.Text += delegate (String t, Boolean Ajouter)
            {
                if (Ajouter)
                    _OngletLog.Texte.AppendText(t + System.Environment.NewLine);
                else
                    _OngletLog.Texte.Text = t;

                _OngletLog.Texte.SelectionStart = _OngletLog.Texte.TextLength;
                _OngletLog.Texte.ScrollToCaret();
                _OngletLog.Refresh();

            };

            WindowLog.Afficher += delegate ()
            {
                _TaskpaneOngletLog.ShowView();
                _OngletLog.Refresh();
            };
        }

        public void SupprimerTaskpane()
        {
            _OngletLog = null;
            _TaskpaneOngletLog.DeleteView();
            _TaskpaneOngletParametres.DeleteView();
            Marshal.ReleaseComObject(_TaskpaneOngletLog);
            Marshal.ReleaseComObject(_TaskpaneOngletParametres);
            _TaskpaneOngletLog = null;
            _TaskpaneOngletParametres = null;
        }
    }
}
