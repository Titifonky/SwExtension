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
        private int indice = 1;

        private String GenRepereDossier
        {
            get
            {
                return "P" + (indice++).ToString();
            }
        }

        protected override void Command()
        {
            try
            {
                ModelDoc2 mdlBase = App.ModelDoc2;
                var ListeCorps = new List<Corps>();

                if (mdlBase.TypeDoc() == eTypeDoc.Piece)
                {
                    foreach (var fDossier in mdlBase.ePartDoc().eListeDesFonctionsDePiecesSoudees())
                    {
                        BodyFolder Dossier = fDossier.GetSpecificFeature2();

                        if (Dossier.IsNull() || Dossier.eNbCorps() == 0 || !(Dossier.eEstUnDossierDeBarres() || Dossier.eEstUnDossierDeToles()))
                            continue;

                        fDossier.Name = "P" + (indice++).ToString();
                    }
                }
                else
                {
                    // On recupère la liste des composants et config
                    // pour pouvoir mettre à jour reinitialiser les noms de dossier
                    var r = mdlBase.eComposantRacine().eDenombrerComposant(
                        comp =>
                        {
                            if (comp.TypeDoc() != eTypeDoc.Piece) return false;

                            foreach (var fDossier in comp.eListeDesFonctionsDePiecesSoudees())
                            {
                                BodyFolder Dossier = fDossier.GetSpecificFeature2();
                                if (Dossier.IsRef() || Dossier.eNbCorps() > 0 || Dossier.eEstUnDossierDeBarres() || Dossier.eEstUnDossierDeToles())
                                    return true;
                            }

                            return false;
                        },
                        // On ne parcourt pas les assemblages exclus
                        comp =>
                        {
                            if (comp.ExcludeFromBOM)
                                return false;

                            return true;
                        }
                        );

                    // On boucle sur les composants
                    // et on modifie les noms de dossier pour
                    // qu'il n'y ai pas de conflits plus tard
                    foreach (var mdl in r.Keys)
                    {
                        mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                        foreach (var cfg in r[mdl].Keys)
                        {
                            mdl.ShowConfiguration2(cfg);
                            mdl.EditRebuild3();
                            foreach (var fDossier in mdl.ePartDoc().eListeDesFonctionsDePiecesSoudees())
                            {
                                BodyFolder Dossier = fDossier.GetSpecificFeature2();
                                if (Dossier.IsRef() || Dossier.eNbCorps() > 0 || Dossier.eEstUnDossierDeBarres() || Dossier.eEstUnDossierDeToles())
                                    fDossier.Name = fDossier.Name + "___Mod";
                            }
                        }

                        if (mdl.GetPathName() != mdlBase.GetPathName())
                            App.Sw.CloseDoc(mdl.GetPathName());
                    }

                    // Comparaison des corps
                    // et synthese des dossiers par corps
                    mdlBase.eComposantRacine().eRecParcourirComposantBase(
                        comp =>
                        {
                            if (comp.TypeDoc() != eTypeDoc.Piece) return;

                            foreach (var fDossier in comp.eListeDesFonctionsDePiecesSoudees())
                            {
                                BodyFolder Dossier = fDossier.GetSpecificFeature2();
                                if (Dossier.IsNull() || Dossier.eNbCorps() == 0 || !(Dossier.eEstUnDossierDeBarres() || Dossier.eEstUnDossierDeToles()))
                                    continue;

                                var SwCorps = Dossier.ePremierCorps();
                                if (SwCorps.IsNull()) continue;

                                var MateriauCorps = SwCorps.eGetMateriauCorpsOuComp(comp);

                                Boolean Ajoute = false;
                                foreach (var CorpsTest in ListeCorps)
                                {
                                    if (MateriauCorps != CorpsTest.Materiau) continue;

                                    if (SwCorps.eEstSemblable(CorpsTest.SwCorps))
                                    {
                                        CorpsTest.Nb += Dossier.eNbCorps();
                                        CorpsTest.AjouterModele(comp, fDossier.Name);

                                        Ajoute = true;
                                        break;
                                    }
                                }

                                if (Ajoute == false)
                                {
                                    var rep = GenRepereDossier;

                                    var corps = new Corps(SwCorps, MateriauCorps);
                                    corps.Nb = Dossier.eNbCorps();
                                    corps.Repere = rep;
                                    corps.AjouterModele(comp, fDossier.Name);
                                    ListeCorps.Add(corps);
                                }
                            }
                        },
                        // On ne parcourt pas les assemblages exclus
                        comp =>
                        {
                            if (comp.ExcludeFromBOM)
                                return false;

                            return true;
                        }
                        );

                    // Petit récap
                    WindowLog.EcrireF("Nb de corps unique : {0}", ListeCorps.Count);
                    int nbtt = 0;
                    foreach (var corps in ListeCorps)
                    {
                        nbtt += corps.Nb;
                        WindowLog.EcrireF("{0} : {1}", corps.Repere, corps.Nb);
                    }

                    WindowLog.EcrireF("Nb total de corps : {0}", nbtt);

                    //==============================
                    // Numerotation des dossier
                    //==============================

                    // On tri les corps par modele et configuration
                    // pour éviter de changer de moele à chaque nouveau corps
                    var ListeModele = new SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<String, String>>>(new CompareModelDoc2());

                    foreach (var corps in ListeCorps)
                    {
                        foreach (var mdl in corps.ListeModele.Keys)
                        {
                            if(!ListeModele.ContainsKey(mdl))
                            {
                                ListeModele.Add(mdl, corps.ListeModele[mdl]);
                            }
                            else
                            {
                                foreach (var cfg in corps.ListeModele[mdl].Keys)
                                {
                                    var lCfg = ListeModele[mdl];
                                    if (!lCfg.ContainsKey(cfg))
                                        lCfg.Add(cfg, corps.ListeModele[mdl][cfg]);
                                    else
                                    {
                                        foreach (var dossier in corps.ListeModele[mdl][cfg].Keys)
                                        {
                                            var lDossier = ListeModele[mdl][cfg];
                                            if (!lDossier.ContainsKey(dossier))
                                                lDossier.Add(dossier, corps.Repere);
                                        }
                                    }
                                }
                                
                            }
                        }
                    }

                    // On renomme les dossiers
                    // en ouvrant les modeles pour que les
                    // modifications soient bien validées
                    var ListeTraite = new HashSet<String>();

                    foreach (var mdl in ListeModele.Keys)
                    {
                        mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                        foreach (var cfg in ListeModele[mdl].Keys)
                        {
                            mdl.ShowConfiguration2(cfg);
                            mdl.EditRebuild3();
                            foreach (var tDossier in ListeModele[mdl][cfg])
                            {
                                // Si le repere n'a jamais été traité,
                                // on vérifie qu'il nexiste pas une fonction avec un nom identique
                                // si oui, on la renomme
                                if (!ListeTraite.Contains(mdl.GetPathName() + "__" + cfg + "__" + tDossier.Value))
                                {
                                    Feature fmod = mdl.ePartDoc().FeatureByName(tDossier.Value);
                                    if (fmod.IsRef())
                                        fmod.Name = fmod.Name + "_mod";
                                }

                                Feature f = mdl.ePartDoc().FeatureByName(tDossier.Key);
                                var nomDossier = NommerDossier(f, tDossier.Value);

                                CustomPropertyManager PM = f.CustomPropertyManager;

                                // Si la propriete existe, on récupère la valeur
                                if (!PM.ePropExiste(CONSTANTES.REF_DOSSIER))
                                {
                                    var ValProp = String.Format("SW-CutListItemName@@@{0}@{1}", nomDossier, mdl.eNomAvecExt());
                                    PM.ePropAdd(CONSTANTES.REF_DOSSIER, ValProp);
                                }

                                ListeTraite.Add(mdl.GetPathName() + "__" + cfg + "__" + tDossier.Value);
                            }
                        }

                        if (mdl.GetPathName() != mdlBase.GetPathName())
                            App.Sw.CloseDoc(mdl.GetPathName());
                    }

                }
            }
            catch (Exception e) { this.LogMethode(new Object[] { e }); }
        }

        // On rajoute des espaces au nom pour remedier au problèmes
        // des corps identiques dans des dossiers différents
        private String NommerDossier(Feature f, String rep)
        {
            int Boucle = 0;
            var rtmp = rep;
            f.Name = rtmp;
            while((f.Name != rtmp) && (Boucle++ < 15))
            {
                rtmp += " ";
                f.Name = rtmp;
            }

            return rtmp;
        }

        private class Corps
        {
            public Body2 SwCorps;
            public String Materiau;
            public String Repere;
            public SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<String, String>>> ListeModele = new SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<String, String>>>(new CompareModelDoc2());
            public int Nb = 0;

            public Corps(Body2 swCorps, String materiau)
            {
                SwCorps = swCorps;
                Materiau = materiau;
            }

            public void AjouterModele(ModelDoc2 mdl, String config, String dossier)
            {
                if(ListeModele.ContainsKey(mdl))
                {
                    var lCfg = ListeModele[mdl];
                    if (lCfg.ContainsKey(config))
                    {
                        var lDossier = lCfg[config];
                        if (!lDossier.ContainsKey(dossier))
                            lDossier.Add(dossier, Repere);
                    }
                    else
                    {
                        var lDossier = new SortedDictionary<String,String>(new WindowsStringComparer());
                        lDossier.Add(dossier, Repere);
                        lCfg.Add(config, lDossier);
                    }
                }
                else
                {
                    var lDossier = new SortedDictionary<String, String>(new WindowsStringComparer());
                    lDossier.Add(dossier, Repere);
                    var lCfg = new SortedDictionary<String, SortedDictionary<String, String>>(new WindowsStringComparer());
                    lCfg.Add(config, lDossier);
                    ListeModele.Add(mdl, lCfg);
                }
            }

            public void AjouterModele(Component2 comp, String dossier)
            {
                AjouterModele(comp.eModelDoc2(), comp.eNomConfiguration(), dossier);
            }
        }
    }
}
