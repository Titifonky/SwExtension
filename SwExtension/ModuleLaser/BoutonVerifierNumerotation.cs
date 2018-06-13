using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;

namespace ModuleLaser
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Verifier la numerotation"),
        ModuleNom("VerifierNumerotation")]

    public class BoutonVerifierNumerotation : BoutonBase
    {
        protected override void Command()
        {
            try
            {
                ModelDoc2 MdlBase = App.ModelDoc2;

                var ListeComposants = MdlBase.ListerComposants(false);

                eTypeCorps Filtre = eTypeCorps.Barre | eTypeCorps.Tole;

                var Dic = new HashSet<String>();

                Boolean Erreur = false;

                Predicate<Feature> Test =
                    f =>
                    {
                        BodyFolder dossier = f.GetSpecificFeature2();
                        if (dossier.IsRef() && dossier.eNbCorps() > 0 && Filtre.HasFlag(dossier.eTypeDeDossier()))
                        {
                            var RefDossier = dossier.eProp(CONSTANTES.REF_DOSSIER);
                            if (String.IsNullOrWhiteSpace(RefDossier))
                                return true;
                        }

                        return false;
                    };

                foreach (var mdl in ListeComposants.Keys)
                {
                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                    foreach (var t in ListeComposants[mdl])
                    {
                        var cfg = t.Key;
                        var nbCfg = t.Value;
                        mdl.ShowConfiguration2(cfg);
                        mdl.EditRebuild3();
                        var Piece = mdl.ePartDoc();

                        foreach (var fDossier in Piece.eListeDesFonctionsDePiecesSoudees(Test))
                        {
                            WindowLog.EcrireF("{0} \"{1}\"", mdl.eNomSansExt(), cfg);
                            WindowLog.EcrireF("  {0} : Pas de reference", fDossier.Name);
                            Erreur = true;
                        }
                    }

                    if (mdl.GetPathName() != MdlBase.GetPathName())
                        App.Sw.CloseDoc(mdl.GetPathName());
                }

                if (!Erreur)
                    WindowLog.Ecrire("Aucune erreur\n");
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }
    }
}
