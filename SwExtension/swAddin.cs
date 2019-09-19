using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

namespace SwExtension
{

    [Guid("F1F96231-4928-4CAC-A81E-CB5B2AB4B3BE"), ComVisible(true)]
    [SwAddin(Description = "Extension des fonctions Sw", Title = "Sw Extension", LoadAtStartup = true)]
    public partial class swAddin : ISwAddin
    {
        private TaskpaneView _TaskpaneOngletLog;
        private OngletLog _OngletLog;

        private TaskpaneView _TaskpaneOngletParametres;
        private OngletParametres _OngletParametres;

        private TaskpaneView _TaskpaneOngletDessin;
        private OngletDessin _OngletDessin;

        bool ISwAddin.ConnectToSW(object ThisSW, int Cookie)
        {
            try
            {
                Log.Demarrer();

                App.Init((SldWorks)ThisSW, this, Cookie);

                App.Sw.SetAddinCallbackInfo2(0, this, App.AddInCookie);

                CreerTaskpane();
                CreerMenusEtOnglets(App.CommandManager);
                AddEventHooks();
                AppliquerOptions();

                return true;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }

            return false;
        }

        private void AppliquerOptions()
        {
            CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator = ".";
        }

        bool ISwAddin.DisconnectFromSW()
        {
            SupprimerTaskpane();
            SupprimerMenus();
            RemoveEventHooks();
            SupprimerPMP();

            // On nettoie les références à SW
            App.Nettoyer();
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
            String CheminImageOngletDessin = Path.Combine(CheminDossier, "Icon_OngletDessin." + ImageFormat.Bmp.ToString().ToLower());

            Image ImgOngletLog = "L".eConvertirEnBmp(18, 16);
            Image ImgOngletParametre = "P".eConvertirEnBmp(18, 16);
            Image ImgOngletDessin = "D".eConvertirEnBmp(18, 16);

            ImgOngletLog.Save(CheminImageOngletLog, ImageFormat.Bmp);
            ImgOngletParametre.Save(CheminImageOngletParametre, ImageFormat.Bmp);
            ImgOngletDessin.Save(CheminImageOngletDessin, ImageFormat.Bmp);
            
            _TaskpaneOngletParametres = App.Sw.CreateTaskpaneView2(CheminImageOngletParametre, "Parametres");
            _TaskpaneOngletDessin = App.Sw.CreateTaskpaneView2(CheminImageOngletDessin, "Dessin");
            _TaskpaneOngletLog = App.Sw.CreateTaskpaneView2(CheminImageOngletLog, "Log");

            _OngletLog = new OngletLog();
            _OngletParametres = new OngletParametres(App.Sw);
            _OngletDessin = new OngletDessin(App.Sw);
            
            _TaskpaneOngletParametres.DisplayWindowFromHandlex64(_OngletParametres.Handle.ToInt64());
            _TaskpaneOngletDessin.DisplayWindowFromHandlex64(_OngletDessin.Handle.ToInt64());
            _TaskpaneOngletLog.DisplayWindowFromHandlex64(_OngletLog.Handle.ToInt64());

            App.Sw.ActiveDocChangeNotify += _OngletParametres.ActiveDocChange;
            App.Sw.ActiveModelDocChangeNotify += _OngletParametres.ActiveDocChange;
            App.Sw.FileCloseNotify += _OngletParametres.CloseDoc;

            App.Sw.ActiveDocChangeNotify += _OngletDessin.ActiveDocChange;
            App.Sw.ActiveModelDocChangeNotify += _OngletDessin.ActiveDocChange;
            App.Sw.FileCloseNotify += _OngletDessin.CloseDoc;

            App.Sw.ActiveDocChangeNotify += _OngletParametres.Rechercher_Propriete_Modele;
            App.Sw.ActiveModelDocChangeNotify += _OngletParametres.Rechercher_Propriete_Modele;
            App.Sw.FileCloseNotify += delegate (String nomFichier, int raison) { return _OngletParametres.Rechercher_Propriete_Modele(); };

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
            _OngletParametres = null;
            _OngletDessin = null;
            _TaskpaneOngletLog.DeleteView();
            _TaskpaneOngletParametres.DeleteView();
            _TaskpaneOngletDessin.DeleteView();
            Marshal.ReleaseComObject(_TaskpaneOngletLog);
            Marshal.ReleaseComObject(_TaskpaneOngletParametres);
            Marshal.ReleaseComObject(_TaskpaneOngletDessin);
            _TaskpaneOngletLog = null;
            _TaskpaneOngletParametres = null;
            _TaskpaneOngletDessin = null;
        }
    }
}
