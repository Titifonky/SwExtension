using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Activer/Desactiver l'aimantation"),
        ModuleNom("ActiverAimantation")]
    public class BoutonActiverAimantation : BoutonBase
    {
        public BoutonActiverAimantation()
        {
            LogToWindowLog = false;
        }

        protected override void Command()
        {
            try
            {
                Frame F = App.Sw.Frame();

                if (App.Sw.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference))
                {
                    App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
                    F.SetStatusBarText("Aimantation désactivée");
                }
                else
                {
                    App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true);
                    F.SetStatusBarText("Aimantation activée");
                }
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }
}
