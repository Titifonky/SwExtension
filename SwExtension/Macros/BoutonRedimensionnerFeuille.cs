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
                var dessin = MdlBase.eDrawingDoc();
                var feuille = dessin.eFeuilleActive(); ;

                dessin.ActivateSheet(feuille.GetName());
                WindowLog.Ecrire("  - " + feuille.GetName());
                feuille.eAjusterAutourDesVues();
                MdlBase.eZoomEtendu();

                //dessin.eParcourirLesFeuilles(
                //    f =>
                //    {
                //        dessin.ActivateSheet(f.GetName());
                //        WindowLog.Ecrire("  - " + f.GetName());
                //        f.eAjusterAutourDesVues();
                //        MdlBase.eZoomEtendu();
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
