using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Mettre en page les feuilles"),
        ModuleNom("MettreEnPage")]
    public class BoutonMettreEnPage : BoutonBase
    {
        protected override void Command()
        {
            DrawingDoc dessin = App.ModelDoc2.eDrawingDoc();

            String FeuilleCourante = ((Sheet)dessin.GetCurrentSheet()).GetName();

            Boolean HauteQualite = true;

            WindowLog.Ecrire("Haute qualité : " + (HauteQualite ? "oui" : "non"));
            WindowLog.SautDeLigne();

            dessin.eParcourirLesFeuilles(
                f =>
                {
                    dessin.ActivateSheet(f.GetName());
                    WindowLog.Ecrire(" - " + f.GetName());
                    String res = dessin.eMettreEnPagePourImpression(f, swPageSetupDrawingColor_e.swPageSetup_AutomaticDrawingColor, HauteQualite);
                    WindowLog.Ecrire("    " + res);
                    WindowLog.SautDeLigne();
                    return false;
                }
                );

            dessin.ActivateSheet(FeuilleCourante);
        }
    }
}
