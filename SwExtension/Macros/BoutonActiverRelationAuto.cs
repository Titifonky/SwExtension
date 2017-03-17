using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Activer/Desactiver les relation auto"),
        ModuleNom("ActiverRelationAuto")]
    public class BoutonActiverRelationAuto : BoutonBase
    {
        public BoutonActiverRelationAuto()
        {
            LogToWindowLog = false;
        }

        protected override void Command()
        {
            try
            {
                Frame F = App.Sw.Frame();

                if (App.Sw.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchAutomaticRelations))
                {
                    App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchAutomaticRelations, false);
                    F.SetStatusBarText("Relation auto désactivée");
                }
                else
                {
                    App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchAutomaticRelations, true);
                    F.SetStatusBarText("Relation auto activée");
                }
            }
            catch (Exception e)
            { Log.Message(e); }
        }
    }
}
