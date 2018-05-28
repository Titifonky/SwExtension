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
                            var LstDossier = c.eListeDesDossiersDePiecesSoudees();
                            foreach (var dossier in LstDossier)
                            {
                                if (!dossier.eEstExclu() && dossier.eEstUnDossierDeToles())
                                {
                                    String Ep = dossier.ePremierCorps().eEpaisseur().ToString();

                                    ListeEp.AddIfNotExist(Ep);
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
                    if (!dossier.eEstExclu() && dossier.eEstUnDossierDeToles())
                    {
                        String Ep = dossier.ePremierCorps().eEpaisseur().ToString();

                        ListeEp.AddIfNotExist(Ep);
                    }
                }
            }

            return ListeEp;
        }

        public static List<string> ListeProfil(this ModelDoc2 mdl)
        {
            List<string> ListeProfil = new List<string>();

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
                                if (!dossier.eEstExclu() && dossier.eEstUnDossierDeToles())
                                {
                                    if (dossier.ePropExiste(CONSTANTES.PROFIL_NOM))
                                    {
                                        String Profil = dossier.eProp(CONSTANTES.PROFIL_NOM);

                                        ListeProfil.AddIfNotExist(Profil);
                                    }
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
                    if (!dossier.eEstExclu() && dossier.eEstUnDossierDeToles())
                    {
                        if (dossier.ePropExiste(CONSTANTES.PROFIL_NOM))
                        {
                            String Profil = dossier.eProp(CONSTANTES.PROFIL_NOM);

                            ListeProfil.AddIfNotExist(Profil);
                        }
                    }
                }
            }

            return ListeProfil;
        }

        /// <summary>
        /// Renvoi la liste unique des modeles et configurations
        /// </summary>
        /// <param name="mdlBase"></param>
        /// <param name="composantsExterne"></param>
        /// <param name="filtreTypeCorps"></param>
        /// <returns></returns>
        public static SortedDictionary<ModelDoc2, SortedDictionary<String, int>> ListerComposants(this ModelDoc2 mdlBase, Boolean composantsExterne, eTypeCorps filtreTypeCorps)
        {
            SortedDictionary<ModelDoc2, SortedDictionary<String, int>> dic = new SortedDictionary<ModelDoc2, SortedDictionary<String, int>>(new CompareModelDoc2());

            if (mdlBase.TypeDoc() == eTypeDoc.Piece)
            {
                var ConfigActive = mdlBase.eNomConfigActive();
                if (!ConfigActive.eEstConfigPliee())
                {
                    WindowLog.Ecrire("Pas de configuration valide," +
                                        "\r\n le nom de la config doit être composée exclusivement de chiffres");
                    return dic;
                }
                var Piece = mdlBase.ePartDoc();
                var lcfg = new SortedDictionary<String, int>(new WindowsStringComparer());
                dic.Add(mdlBase, lcfg);
                foreach (var dossier in Piece.eListeDesDossiersDePiecesSoudees())
                {
                    if (!dossier.eEstExclu() &&
                        filtreTypeCorps.HasFlag(dossier.eTypeDeDossier()))
                    {
                        lcfg.Add(ConfigActive, 1);
                        break;
                    }
                }
            }
            else
            {
                mdlBase.eComposantRacine().eRecParcourirComposantBase(
                        comp =>
                        {
                            if (comp.IsSuppressed() || comp.ExcludeFromBOM) return;

                            var mdl = comp.eModelDoc2();
                            var cfg = comp.eNomConfiguration();
                            if (dic.ContainsKey(mdl))
                                if (dic[mdl].ContainsKey(cfg))
                                {
                                    dic[mdl][cfg] += 1;
                                    return;
                                }

                            foreach (var fDossier in comp.eListeDesFonctionsDePiecesSoudees())
                            {
                                BodyFolder SwDossier = fDossier.GetSpecificFeature2();
                                if (SwDossier.IsRef() ||
                                SwDossier.eNbCorps() > 0 ||
                                !SwDossier.eEstExclu() ||
                                filtreTypeCorps.HasFlag(SwDossier.eTypeDeDossier()))
                                {
                                    if (dic.ContainsKey(mdl))
                                        dic[mdl].Add(cfg, 1);
                                    else
                                    {
                                        var lcfg = new SortedDictionary<String, int>(new WindowsStringComparer());
                                        lcfg.Add(cfg, 1);
                                        dic.Add(mdl, lcfg);
                                    }
                                    break;
                                }
                            }
                        },
                        // On ne parcourt pas les assemblages exclus
                        c =>
                        {
                            if (c.ExcludeFromBOM)
                                return false;

                            return true;
                        }
                        );
            }

            return dic;
        }

        /// <summary>
        /// Renvoi la liste des modeles avec la liste des configurations, des dossiers
        /// et les quantites de chaque dossier dans l'assemblage
        /// </summary>
        /// <param name="mdlBase"></param>
        /// <param name="composantsExterne"></param>
        /// <param name="filtreDossier"></param>
        /// <returns></returns>
        public static SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<int, int>>> DenombrerDossiers(this ModelDoc2 mdlBase, Boolean composantsExterne, Predicate<Feature> filtreDossier = null, Predicate<Component2> filtreComposant = null)
        {
            SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<int, int>>> dic = new SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<int, int>>>(new CompareModelDoc2());

            if (mdlBase.TypeDoc() == eTypeDoc.Piece)
            {
                var ConfigActive = mdlBase.eNomConfigActive();
                if (!ConfigActive.eEstConfigPliee())
                {
                    WindowLog.Ecrire("Pas de configuration valide," +
                                        "\r\n le nom de la config doit être composée exclusivement de chiffres");
                    return dic;
                }
                var Piece = mdlBase.ePartDoc();
                var dicDossier = new SortedDictionary<int, int>();
                foreach (var dossier in Piece.eListeDesDossiersDePiecesSoudees())
                {
                    if (dossier.eNbCorps() > 0 && !dossier.eEstExclu() && (filtreDossier.IsNull() || filtreDossier(dossier.GetFeature())))
                    {
                        dicDossier.Add(dossier.GetFeature().GetID(), dossier.eNbCorps());
                    }
                }

                var dicConfig = new SortedDictionary<String, SortedDictionary<int, int>>(new WindowsStringComparer());
                dicConfig.Add(ConfigActive, dicDossier);
                dic.Add(mdlBase, dicConfig);
            }
            else
            {
                var ListeDossiers = new List<Dossier>();

                mdlBase.eComposantRacine().eRecParcourirComposantBase(
                        comp =>
                        {
                            if (comp.IsSuppressed() || comp.ExcludeFromBOM) return;

                            foreach (var fDossier in comp.eListeDesFonctionsDePiecesSoudees())
                            {
                                BodyFolder SwDossier = fDossier.GetSpecificFeature2();
                                if (SwDossier.IsRef() && SwDossier.eNbCorps() > 0 && !SwDossier.eEstExclu() && (filtreDossier.IsNull() || filtreDossier(fDossier)))
                                {
                                    Boolean Ajoute = false;
                                    foreach (var DossierTest in ListeDossiers)
                                    {
                                        var RefDossier = SwDossier.eProp(CONSTANTES.REF_DOSSIER);
                                        if (RefDossier == DossierTest.Repere)
                                        {
                                            if (filtreComposant.IsRef() && filtreComposant(comp))
                                                DossierTest.FiltreComposant = true;

                                            DossierTest.Nb += SwDossier.eNbCorps();
                                            Ajoute = true;
                                            break;
                                        }
                                    }

                                    if (Ajoute == false)
                                    {
                                        var RefDossier = SwDossier.eProp(CONSTANTES.REF_DOSSIER);
                                        var dossier = new Dossier(RefDossier, comp.eModelDoc2(), comp.eNomConfiguration(), fDossier.GetID());

                                        // Si le filtre renvoi false
                                        // on desactive la propriete filtre
                                        if (filtreComposant.IsRef() && !filtreComposant(comp))
                                            dossier.FiltreComposant = false;

                                        dossier.Nb = SwDossier.eNbCorps();
                                        ListeDossiers.Add(dossier);
                                    }
                                }
                            }
                        },
                        // On ne parcourt pas les assemblages exclus
                        c =>
                        {
                            if (c.ExcludeFromBOM)
                                return false;

                            return true;
                        }
                        );

                foreach (var dossier in ListeDossiers)
                {
                    if (!dossier.FiltreComposant)
                        continue;

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

            return dic;
        }

        private class Dossier
        {
            public int Id;
            public String Repere;
            public ModelDoc2 Mdl;
            public String Config;
            public int Nb = 0;
            public Boolean FiltreComposant = true;

            public Dossier(String repere, ModelDoc2 mdl, String config, int id)
            {
                Repere = repere;
                Mdl = mdl;
                Config = config;
                Id = id;
            }
        }
    }
}
