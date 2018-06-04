using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Redimensionner la feuille"),
        ModuleNom("RedimensionnerFeuille")]
    public class BoutonRedimensionnerFeuille : BoutonBase
    {
        protected override void Command()
        {
            try
            {
                var dessin = App.ModelDoc2.eDrawingDoc();
                var feuille = dessin.eFeuilleActive(); ;

                dessin.ActivateSheet(feuille.GetName());
                WindowLog.Ecrire("  - " + feuille.GetName());
                feuille.eAjusterAutourDesVues();
                App.ModelDoc2.eZoomEtendu();

                //dessin.eParcourirLesFeuilles(
                //    f =>
                //    {
                //        dessin.ActivateSheet(f.GetName());
                //        WindowLog.Ecrire("  - " + f.GetName());
                //        f.eAjusterAutourDesVues();
                //        App.ModelDoc2.eZoomEtendu();
                //        return false;
                //    }
                //    );
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }
}
