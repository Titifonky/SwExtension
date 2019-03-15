using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Mettre en page les feuilles"),
        ModuleNom("MettreEnPage")]

    public class BoutonMettreEnPage : BoutonBase
    {
        protected override void Command()
        {
            try
            {
                DrawingDoc dessin = MdlBase.eDrawingDoc();

                String FeuilleCourante = dessin.eFeuilleActive().GetName();

                Boolean HauteQualite = true;

                WindowLog.Ecrire("Haute qualité : " + (HauteQualite ? "oui" : "non"));
                WindowLog.SautDeLigne();

                MdlBase.Extension.UsePageSetup = (int)swPageSetupInUse_e.swPageSetupInUse_DrawingSheet;

                App.Sw.RunCommand((int)swCommands_e.swCommands_Page_Setup, "");

                dessin.eParcourirLesFeuilles(
                    f =>
                    {
                        //dessin.ActivateSheet(f.GetName());
                        WindowLog.Ecrire(" - " + f.GetName());
                        String res = dessin.eMettreEnPagePourImpression(f, swPageSetupDrawingColor_e.swPageSetup_AutomaticDrawingColor, HauteQualite);
                        WindowLog.Ecrire("    " + res);
                        WindowLog.SautDeLigne();
                        return false;
                    }
                    );

                //dessin.ActivateSheet(FeuilleCourante);
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }
}
