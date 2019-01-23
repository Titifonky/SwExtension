using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;

namespace ModuleProduction
{
    [ModuleTypeDocContexte(eTypeDoc.Piece),
        ModuleTitre("Afficher/Masquer Esquisse reperage"),
        ModuleNom("AfficherMasquerEsquisseReperage")]

    public class BoutonAfficherMasquerEsquisseReperage : BoutonBase
    {
        public BoutonAfficherMasquerEsquisseReperage() { }

        protected override void Command()
        {
            try
            {
                Feature Esquisse = MdlBase.pEsquisseRepere(false);

                Esquisse.eSelect();
                MdlBase.UnblankSketch();
                Esquisse.SetUIState((int)swUIStates_e.swIsHiddenInFeatureMgr, false);

                MdlBase.eEffacerSelection();
                MdlBase.EditRebuild3();

            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }
}
