using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using SwExtension.ModuleEchelleVue;
using System;

namespace ModuleEchelleVue
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Echelle de la vue"),
        ModuleNom("EchelleVue")]
    public class BoutonEchelleVue : BoutonBase
    {
        private Parametre PositionVue;

        public BoutonEchelleVue()
        {
            PositionVue = _Config.AjouterParam("PositionVue", "0:0", "Position de la vue :");

            LogToWindowLog = false;
        }

        protected override void Command()
        {
            try
            {
                if(InfoFenetre.Form.IsNull())
                {
                    Frame f = (Frame)App.Sw.Frame();
                    
                    InfoFenetre.Form = new FormEchelleVue(MdlBase.eDrawingDoc(), PositionVue, _Config);
                    InfoFenetre.Form.Show(new WindowWrapper((IntPtr)f.GetHWnd()));
                }
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }

    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        private IntPtr _hwnd;
        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }
        public IntPtr Handle
        {
            get { return _hwnd; }
        }
    }


    public static class InfoFenetre
    {
        public static FormEchelleVue Form = null;
    }  
}
