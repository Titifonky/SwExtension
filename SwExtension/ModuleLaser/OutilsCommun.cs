using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;

namespace ModuleLaser
{
    public static class OutilsCommun
    {
        private const String ATTRIBUT_INDEX_DOSSIER = "NoDossierPiecesSoudees";
        private const String ATTRIBUT_PARAM_INDEXMAX = "Index";

        private static AttributeDef AttDef = null;

        public static Parameter eAttributNoDossier(this Component2 cp)
        {
            return cp.ePartDoc().eAttributNoDossier();
        }

        /// <summary>
        /// Retourne l'attribut NoDossier
        /// </summary>
        /// <param name="piece"></param>
        /// <returns></returns>
        public static Parameter eAttributNoDossier(this PartDoc piece)
        {
            if (AttDef.IsNull())
            {
                AttDef = App.Sw.DefineAttribute(ATTRIBUT_INDEX_DOSSIER);
                AttDef.AddParameter(ATTRIBUT_PARAM_INDEXMAX, (int)swParamType_e.swParamTypeInteger, 0, 0);
                AttDef.Register();
            }

            // Recherche de l'attribut dans la piece
            SolidWorks.Interop.sldworks.Attribute Att = null;
            ModelDoc2 mdl = piece.eModelDoc2();
            Parameter P = null;
            Feature F = mdl.eChercherFonction(f => { return f.Name == ATTRIBUT_INDEX_DOSSIER; });
            if (F.IsRef())
            {
                Att = F.GetSpecificFeature2();

                P = (Parameter)Att.GetParameter(ATTRIBUT_PARAM_INDEXMAX);

                if (P.IsNull())
                {
                    Att.Delete(false);
                    Att = null;
                }
            }

            if (Att.IsNull())
            {
                Att = AttDef.CreateInstance5(mdl, null, ATTRIBUT_INDEX_DOSSIER, 1, (int)swInConfigurationOpts_e.swAllConfiguration);
                P = (Parameter)Att.GetParameter(ATTRIBUT_PARAM_INDEXMAX);
            }

            return P;
        }

        /// <summary>
        /// Met l'attribut NoDossier à 0
        /// </summary>
        /// <param name="piece"></param>
        public static void eReinitialiserNoDossierMax(this PartDoc piece)
        {
            // On récupère la valeur du parametre
            Parameter P = eAttributNoDossier(piece);
            P.SetDoubleValue2(0, (int)swInConfigurationOpts_e.swAllConfiguration, "");
        }

        /// <summary>
        /// Met l'attribut NoDossier à 0
        /// </summary>
        /// <param name="cp"></param>
        public static void eReinitialiserNoDossierMax(this Component2 cp)
        {
            cp.ePartDoc().eReinitialiserNoDossierMax();
        }

        /// <summary>
        /// Numerote les dossier à partir de noDossierDepart
        /// </summary>
        /// <param name="comp"></param>
        /// <param name="MajListeDesDossiersDePiecesSoudees"></param>
        /// <returns>Retourne la liste des dossiers et le no</returns>
        public static Dictionary<String, int> eNumeroterDossier(this Component2 comp, Boolean MajListeDesDossiersDePiecesSoudees = false)
        {
            if (MajListeDesDossiersDePiecesSoudees)
                comp.eMajListeDesPiecesSoudees();

            List<BodyFolder> listeDossier = comp.eListeDesDossiersDePiecesSoudees();
            Parameter Param = eAttributNoDossier(comp);

            return eFonctionNumeroterDossier(Param, listeDossier);
        }

        /// <summary>
        /// Numerote les dossier à partir de noDossierDepart
        /// </summary>
        /// <param name="piece"></param>
        /// <param name="MajListeDesDossiersDePiecesSoudees"></param>
        /// <returns>Retourne la liste des dossiers et le no</returns>
        public static Dictionary<String, int> eNumeroterDossier(this PartDoc piece, Boolean MajListeDesDossiersDePiecesSoudees = false)
        {
            if (MajListeDesDossiersDePiecesSoudees)
                piece.eMajListeDesPiecesSoudees();

            List<BodyFolder> listeDossier = piece.eListeDesDossiersDePiecesSoudees();
            Parameter Param = eAttributNoDossier(piece);

            return eFonctionNumeroterDossier(Param, listeDossier);
        }

        private static Dictionary<String, int> eFonctionNumeroterDossier(Parameter P, List<BodyFolder> listeDossier)
        {
            Dictionary<String, int> Dic = new Dictionary<string, int>();
            int pNoDossier = 0;
            int pNoDossierMax = (int)P.GetDoubleValue();

            foreach (BodyFolder dossier in listeDossier)
            {
                if (dossier.GetBodyCount() == 0) continue;

                CustomPropertyManager PM = dossier.GetFeature().CustomPropertyManager;

                // Si la propriete existe, on récupère la valeur
                if (PM.ePropExiste(CONSTANTES.NO_DOSSIER))
                {
                    String val, result; Boolean wasResolved;
                    PM.Get5(CONSTANTES.NO_DOSSIER, false, out val, out result, out wasResolved);
                    pNoDossier = result.eToInteger();
                    pNoDossierMax = Math.Max(pNoDossierMax, pNoDossier);
                }
                else
                {
                    pNoDossier = ++pNoDossierMax;
                    PM.ePropAdd(CONSTANTES.NO_DOSSIER, pNoDossier);
                }

                Dic.AddIfNotExist(dossier.eNom(), pNoDossier);
            }

            // On met à jour le parametre
            P.SetDoubleValue2(pNoDossierMax, (int)swInConfigurationOpts_e.swAllConfiguration, "");

            return Dic;
        }

        public static void eEffacerNoDossier(this Component2 cp)
        {
            cp.eReinitialiserNoDossierMax();

            var liste = cp.eListeDesDossiersDePiecesSoudees();
            if (liste.IsNullOrEmpty()) return;

            foreach (BodyFolder dossier in liste)
            {
                if (dossier.GetBodyCount() == 0) continue;

                CustomPropertyManager PM = dossier.GetFeature().CustomPropertyManager;

                PM.Delete2(CONSTANTES.NO_DOSSIER);
            }
        }

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
                                    String Materiau = dossier.eGetMateriau();

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

                //foreach (Body2 corps in mdl.ePartDoc().eListeCorps(true))
                //{
                //    if (TypeCorps.HasFlag(corps.eTypeDeCorps()))
                //    {
                //        String Materiau = corps.eGetMateriauCorpsOuPiece(mdl.ePartDoc(), mdl.eNomConfigActive());

                //        ListeMateriaux.AddIfNotExist(Materiau);
                //    }
                //}
            }

            return ListeMateriaux;
        }
    }
}
