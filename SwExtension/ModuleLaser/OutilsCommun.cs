using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ModuleLaser
{
    public static class OutilsCommun
    {
        public static Configuration CreerConfigDepliee(this ModelDoc2 mdl, String NomConfigDepliee, String NomConfigPliee)
        {
            return mdl.ConfigurationManager.AddConfiguration(NomConfigDepliee, NomConfigDepliee, "", 0, NomConfigPliee, "");
        }

        public static String RefPiece(this ModelDoc2 mdl, String Pattern, String configPliee, String noDossier)
        {
            //<Nom_Piece>-<Nom_Config>-<No_Dossier>

            String result = Pattern;

            result = result.Replace("<Nom_Piece>", mdl.eNomSansExt());
            result = result.Replace("<Nom_Config>", configPliee);
            result = result.Replace("<No_Dossier>", noDossier);
            return result;
        }

        public static void DeplierTole(this Body2 Tole, ModelDoc2 mdl, String nomConfigDepliee)
        {
            Feature FonctionDepliee = Tole.eFonctionEtatDepliee();
            FonctionDepliee.eModifierEtat(swFeatureSuppressionAction_e.swUnSuppressFeature, nomConfigDepliee);
            FonctionDepliee.eModifierEtat(swFeatureSuppressionAction_e.swUnSuppressDependent, nomConfigDepliee);

            mdl.eEffacerSelection();

            FonctionDepliee.eParcourirSousFonction(
                f =>
                {
                    f.eModifierEtat(swFeatureSuppressionAction_e.swUnSuppressFeature, nomConfigDepliee);

                    if ((f.Name.ToLowerInvariant() == CONSTANTES.LIGNES_DE_PLIAGE.ToLowerInvariant()) ||
                        (f.Name.ToLowerInvariant() == CONSTANTES.CUBE_DE_VISUALISATION.ToLowerInvariant()))
                    {
                        f.eSelect(true);
                    }
                    return false;
                }
                );

            mdl.UnblankSketch();
            mdl.eEffacerSelection();

            //// Si des corps autre que la tole dépliée sont encore visible dans la config, on les cache et on recontruit tout
            //foreach (Body2 pCorps in mdl.ePartDoc().eListeCorps(false))
            //{
            //    if ((pCorps.Name == CONSTANTES.NOM_CORPS_DEPLIEE))
            //        pCorps.eVisible(true);
            //    else
            //        pCorps.eVisible(false);
            //}

        }

        public static void PlierTole(this Body2 Tole, ModelDoc2 mdl, String nomConfigPliee)
        {
            var tmpTole = Tole;

            if (tmpTole.IsNull())
            {
                tmpTole = mdl.ePartDoc().eChercherCorps(CONSTANTES.NOM_CORPS_DEPLIEE, false);

                if (tmpTole.IsNull()) return;
            }

            Feature FonctionDepliee = tmpTole.eFonctionEtatDepliee();

            FonctionDepliee.eModifierEtat(swFeatureSuppressionAction_e.swSuppressFeature, nomConfigPliee);

            mdl.eEffacerSelection();

            FonctionDepliee.eParcourirSousFonction(
                f =>
                {
                    if ((f.Name.ToLowerInvariant() == CONSTANTES.LIGNES_DE_PLIAGE.ToLowerInvariant()) ||
                        (f.Name.ToLowerInvariant() == CONSTANTES.CUBE_DE_VISUALISATION.ToLowerInvariant()))
                    {
                        f.eSelect(true);
                    }
                    return false;
                }
                );

            mdl.BlankSketch();
            mdl.eEffacerSelection();
        }

        public static List<string> ListeMateriaux(this ModelDoc2 mdl, eTypeCorps TypeCorps)
        {
            List<string> ListeMateriaux = new List<string>();

            if (mdl.TypeDoc() == eTypeDoc.Assemblage)
            {

                App.ModelDoc2.eRecParcourirComposants(
                    c =>
                    {
                        if (!c.IsHidden(false) && !c.ExcludeFromBOM && (c.TypeDoc() == eTypeDoc.Piece))
                        {
                            var LstDossier = c.eListeDesDossiersDePiecesSoudees();
                            foreach (var dossier in LstDossier)
                            {
                                if (!dossier.eEstExclu() && TypeCorps.HasFlag(dossier.eTypeDeDossier()))
                                {
                                    //String Materiau = dossier.eGetMateriau();
                                    String Materiau = dossier.ePremierCorps().eGetMateriauCorpsOuComp(c);

                                    ListeMateriaux.AddIfNotExist(Materiau);
                                }
                            }
                        }

                        return false;
                    }
                );
            }
            else if (mdl.TypeDoc() == eTypeDoc.Piece)
            {
                var LstDossier = mdl.ePartDoc().eListeDesDossiersDePiecesSoudees();
                foreach (var dossier in LstDossier)
                {
                    if (!dossier.eEstExclu() && TypeCorps.HasFlag(dossier.eTypeDeDossier()))
                    {
                        String Materiau = dossier.eGetMateriau();

                        ListeMateriaux.AddIfNotExist(Materiau);
                    }
                }
            }

            return ListeMateriaux;
        }

        public static List<string> ListeEp(this ModelDoc2 mdl)
        {
            List<string> ListeEp = new List<string>();

            if (mdl.TypeDoc() == eTypeDoc.Assemblage)
            {

                App.ModelDoc2.eRecParcourirComposants(
                    c =>
                    {
                        if (!c.IsHidden(false) && !c.ExcludeFromBOM && (c.TypeDoc() == eTypeDoc.Piece))
                        {
                            foreach (var corps in c.eListeCorps())
                            {
                                if(corps.eTypeDeCorps() == eTypeCorps.Tole)
                                    ListeEp.AddIfNotExist(corps.eEpaisseurCorps().ToString());
                            }
                        }

                        return false;
                    }
                );
            }
            else if (mdl.TypeDoc() == eTypeDoc.Piece)
            {
                var LstDossier = mdl.ePartDoc().eListeDesDossiersDePiecesSoudees();
                foreach (var dossier in LstDossier)
                {
                    if (!dossier.eEstExclu() && dossier.eEstUnDossierDeToles())
                    {
                        Double Ep = dossier.eEpaisseur1ErCorpsOuDossier();

                        // On laisse les epaisseurs négatives pour pouvoir les localiser.
                        //if (Ep == -1)
                        //    continue;

                        ListeEp.AddIfNotExist(Ep.ToString());
                    }
                }
            }

            ListeEp.Sort(new WindowsStringComparer());

            return ListeEp;
        }

        /// <summary>
        /// Renvoi la liste unique des modeles et configurations
        /// Modele : ModelDoc2
        ///     |-Config1 : Nom de la configuration
        ///     |     |-Nb : quantite de configuration identique dans le modele complet
        ///     |-Config 2
        ///     | etc...
        /// </summary>
        /// <param name="mdlBase"></param>
        /// <param name="composantsExterne"></param>
        /// <param name="filtreTypeCorps"></param>
        /// <returns></returns>
        public static SortedDictionary<ModelDoc2, SortedDictionary<String, int>> ListerComposants(this ModelDoc2 mdlBase, Boolean composantsExterne)
        {
            SortedDictionary<ModelDoc2, SortedDictionary<String, int>> dic = new SortedDictionary<ModelDoc2, SortedDictionary<String, int>>(new CompareModelDoc2());

            try
            {
                Action<Component2> Test = delegate (Component2 comp)
                {
                    var cfg = comp.eNomConfiguration();

                    if (comp.IsSuppressed() || comp.ExcludeFromBOM || !cfg.eEstConfigPliee() || comp.TypeDoc() != eTypeDoc.Piece) return;

                    var mdl = comp.eModelDoc2();
                    
                    if (dic.ContainsKey(mdl))
                        if (dic[mdl].ContainsKey(cfg))
                        {
                            dic[mdl][cfg] += 1;
                            return;
                        }

                    if (dic.ContainsKey(mdl))
                        dic[mdl].Add(cfg, 1);
                    else
                    {
                        var lcfg = new SortedDictionary<String, int>(new WindowsStringComparer());
                        lcfg.Add(cfg, 1);
                        dic.Add(mdl, lcfg);
                    }
                };

                if (mdlBase.TypeDoc() == eTypeDoc.Piece)
                    Test(mdlBase.eComposantRacine());
                else
                {
                    mdlBase.eComposantRacine().eRecParcourirComposantBase(
                        Test,
                        // On ne parcourt pas les assemblages exclus
                        c =>
                        {
                            if (c.ExcludeFromBOM)
                                return false;

                            return true;
                        }
                        );
                }
            }
            catch (Exception e) { Log.LogErreur(new Object[] { e }); }

            return dic;
        }

        /// <summary>
        /// Renvoi la liste des modeles avec la liste des configurations, des dossiers
        /// et les quantites de chaque dossier dans l'assemblage
        /// Modele : ModelDoc2
        ///     |-Config1 : Nom de la configuration
        ///     |     |-Dossier1 : Comparaison avec la propriete RefDossier, référence à l'Id de la fonction pour pouvoir le selectionner plus tard
        ///     |     |      |- Nb : quantite de corps identique dans le modele complet
        ///     |     |
        ///     |     |-Dossier2
        ///     |            |- Nb
        ///     |-Config 2
        ///     | etc...
        /// </summary>
        /// <param name="mdlBase"></param>
        /// <param name="composantsExterne"></param>
        /// <param name="filtreDossier"></param>
        /// <returns></returns>
        public static SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<int, int>>> DenombrerDossiers(this ModelDoc2 mdlBase, Boolean composantsExterne, Predicate<Feature> filtreDossier = null, Boolean fermerFichier = false)
        {

            SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<int, int>>> dic = new SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<int, int>>>(new CompareModelDoc2());
            try
            {
                var ListeComposants = mdlBase.ListerComposants(composantsExterne);

                var ListeDossiers = new Dictionary<String, Dossier>();

                Predicate<Feature> Test = delegate (Feature fDossier)
                {
                    BodyFolder SwDossier = fDossier.GetSpecificFeature2();
                    if (SwDossier.IsRef() && SwDossier.eNbCorps() > 0 && !SwDossier.eEstExclu() && (filtreDossier.IsNull() || filtreDossier(fDossier)))
                        return true;

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
                            BodyFolder SwDossier = fDossier.GetSpecificFeature2();
                            var RefDossier = SwDossier.eProp(CONSTANTES.REF_DOSSIER);

                            if (ListeDossiers.ContainsKey(RefDossier))
                                ListeDossiers[RefDossier].Nb += SwDossier.eNbCorps() * nbCfg;
                            else
                            {
                                var dossier = new Dossier(RefDossier, mdl, cfg, fDossier.GetID());
                                dossier.Nb = SwDossier.eNbCorps() * nbCfg;

                                ListeDossiers.Add(RefDossier, dossier);
                            }
                        }
                    }

                    if (fermerFichier)
                        mdl.eFermerSiDifferent(mdlBase); ;
                }

                // Conversion d'une liste de dossier
                // en liste de modele
                foreach (var dossier in ListeDossiers.Values)
                {
                    if (dic.ContainsKey(dossier.Mdl))
                    {
                        var lcfg = dic[dossier.Mdl];
                        if (lcfg.ContainsKey(dossier.Config))
                        {
                            var ldossier = lcfg[dossier.Config];
                            ldossier.Add(dossier.Id, dossier.Nb);
                        }
                        else
                        {
                            var ldossier = new SortedDictionary<int, int>();
                            ldossier.Add(dossier.Id, dossier.Nb);
                            lcfg.Add(dossier.Config, ldossier);
                        }
                    }
                    else
                    {
                        var ldossier = new SortedDictionary<int, int>();
                        ldossier.Add(dossier.Id, dossier.Nb);
                        var lcfg = new SortedDictionary<String, SortedDictionary<int, int>>(new WindowsStringComparer());
                        lcfg.Add(dossier.Config, ldossier);
                        dic.Add(dossier.Mdl, lcfg);
                    }
                }
            }
            catch (Exception e) { Log.LogErreur(new Object[] { e }); }

            return dic;
        }

        private class Dossier
        {
            public int Id;
            public String Repere;
            public ModelDoc2 Mdl;
            public String Config;
            public int Nb = 0;

            public Dossier(String repere, ModelDoc2 mdl, String config, int id)
            {
                Repere = repere;
                Mdl = mdl;
                Config = config;
                Id = id;
            }
        }

        private static readonly String ChaineIndice = "ZYXWVUTSRQPONMLKJIHGFEDCBA";

        public static String ChercherIndice(List<String> liste)
        {
            for (int i = 0; i < ChaineIndice.Length; i++)
            {
                if (liste.Any(d => { return d.EndsWith(" Ind " + ChaineIndice[i]) ? true : false; }))
                    return "Ind " + ChaineIndice[Math.Max(0, i - 1)];
            }

            return "Ind " + ChaineIndice.Last();
        }
    }
}
