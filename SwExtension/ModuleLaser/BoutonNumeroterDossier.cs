using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Numeroter les dossiers"),
        ModuleNom("NumeroterDossier")]
    public class BoutonNumeroterDossier : BoutonBase
    {
        protected override void Command()
        {

            try
            {
                ModelDoc2 mdlBase = App.ModelDoc2;
                var ListeComp = mdlBase.eComposantRacine().eRecListeComposant(c => { return c.TypeDoc() == eTypeDoc.Piece; }, null, true);

                foreach (var comp in ListeComp)
                {
                    var ListefDossier = comp.eListeDesFonctionsDePiecesSoudees();

                    for (int i = 0; i < ListefDossier.Count; i++)
                    {
                        Feature fDossier = ListefDossier[i];
                        BodyFolder Dossier = fDossier.GetSpecificFeature2();
                        if (Dossier.IsNull() || Dossier.eNbCorps() == 0 || Dossier.eEstExclu() || !(Dossier.eEstUnDossierDeBarres() || Dossier.eEstUnDossierDeToles()))
                            continue;

                        CustomPropertyManager PM = Dossier.GetFeature().CustomPropertyManager;
                        PM.Delete2(CONSTANTES.NO_DOSSIER);
                    }
                }

                int noDossierMax = 1;

                HashSet<String> DicDossier = new HashSet<string>();

                foreach (var comp in ListeComp)
                {
                    WindowLog.EcrireF("{0}", comp.Name2);

                    ModelDoc2 mdl = comp.eModelDoc2();
                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                    var ListefDossier = mdl.ePartDoc().eListeDesFonctionsDePiecesSoudees();

                    for (int i = 0; i < ListefDossier.Count; i++)
                    {
                        Feature fDossier = ListefDossier[i];
                        BodyFolder Dossier = fDossier.GetSpecificFeature2();
                        if (Dossier.IsNull() || Dossier.eNbCorps() == 0 || Dossier.eEstExclu() || !(Dossier.eEstUnDossierDeBarres() || Dossier.eEstUnDossierDeToles()))
                            continue;

                        CustomPropertyManager PM = Dossier.GetFeature().CustomPropertyManager;
                        var ClefDossier = comp.eNomSansExt() + "-" + fDossier.Name;

                        // Si la propriete existe, on récupère la valeur
                        if (!DicDossier.Contains(ClefDossier))
                        {
                            WindowLog.EcrireF("  {0} -> {1}", fDossier.Name, noDossierMax);
                            PM.ePropAdd(CONSTANTES.NO_DOSSIER, noDossierMax++);
                            DicDossier.Add(ClefDossier);
                        }  
                    }

                    if (comp.GetPathName() != mdlBase.GetPathName())
                        App.Sw.CloseDoc(mdl.GetPathName());
                }
            }
            catch (Exception e) { this.LogMethode(new Object[] { e }); }
        }
    }
}
