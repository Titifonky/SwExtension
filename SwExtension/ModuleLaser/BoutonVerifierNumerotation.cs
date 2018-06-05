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

                eTypeCorps Filtre = eTypeCorps.Barre | eTypeCorps.Tole;

                var Dic = new HashSet<String>();

                var dic = MdlBase.eRecParcourirComposants(
                    comp =>
                    {
                        if (!comp.IsSuppressed())
                        {
                            var clef = comp.eNomAvecExt() + "___" + comp.eNomConfiguration();
                            if (!Dic.Contains(clef))
                            {
                                Dic.Add(clef);

                                var l = comp.eListeDesFonctionsDePiecesSoudees(
                                    f =>
                                    {
                                        BodyFolder dossier = f.GetSpecificFeature2();
                                        if (dossier.IsRef() && dossier.eNbCorps() > 0 && Filtre.HasFlag(dossier.eTypeDeDossier()))
                                        {
                                            var RefDossier = dossier.eProp(CONSTANTES.REF_DOSSIER);
                                            if (String.IsNullOrWhiteSpace(RefDossier))
                                            {
                                                WindowLog.EcrireF("{0} \"{1}\"", comp.eNomSansExt(), comp.eNomConfiguration());
                                                WindowLog.EcrireF("  {0} : Pas de reference", f.Name);
                                            }
                                        }

                                        return true;
                                    }
                                    );
                            }

                        }
                        return false;
                    }
                    );
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }
    }
}
