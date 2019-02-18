using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ModuleProduction
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Nettoyer le reperage"),
        ModuleNom("NettoyerReperage")]

    public class BoutonNettoyerReperage : BoutonBase
    {
        protected override void Command()
        {
            try
            {
                Nettoyer();
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }

        private void SupprimerDefBloc(ModelDoc2 mdl, String cheminbloc)
        {
            var TabDef = (Object[])mdl.SketchManager.GetSketchBlockDefinitions();
            if (TabDef.IsRef())
            {
                foreach (SketchBlockDefinition blocdef in TabDef)
                {
                    if (blocdef.FileName == cheminbloc)
                    {
                        Feature d = blocdef.GetFeature();
                        d.eSelect();
                        mdl.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                        mdl.eEffacerSelection();
                        break;
                    }
                }
            }
        }

        private String CheminBlocEsquisseNumeroter()
        {
            return Sw.CheminBloc(CONSTANTES.NOM_BLOCK_ESQUISSE_NUMEROTER);
        }

        private Feature EsquisseRepere(ModelDoc2 mdl, Boolean creer = true)
        {
            // On recherche l'esquisse contenant les parametres
            Feature Esquisse = mdl.eChercherFonction(fc => { return fc.Name == CONSTANTES.NOM_ESQUISSE_NUMEROTER; });

            if (Esquisse.IsNull() && creer)
            {
                var SM = mdl.SketchManager;

                // On recherche le chemin du bloc
                String cheminbloc = CheminBlocEsquisseNumeroter();

                if (String.IsNullOrWhiteSpace(cheminbloc))
                    return null;

                // On supprime la definition du bloc
                SupprimerDefBloc(mdl, cheminbloc);

                // On recherche le plan de dessus, le deuxième dans la liste des plans de référence
                Feature Plan = mdl.eListeFonctions(fc => { return fc.GetTypeName2() == FeatureType.swTnRefPlane; })[1];

                // Selection du plan et création de l'esquisse
                Plan.eSelect();
                SM.InsertSketch(true);
                SM.AddToDB = false;
                SM.DisplayWhenAdded = true;

                mdl.eEffacerSelection();

                // On récupère la fonction de l'esquisse
                Esquisse = mdl.Extension.GetLastFeatureAdded();

                // On insère le bloc
                MathUtility Mu = App.Sw.GetMathUtility();
                MathPoint Origine = Mu.CreatePoint(new double[] { 0, 0, 0 });
                var def = SM.MakeSketchBlockFromFile(Origine, cheminbloc, false, 1, 0);

                // On récupère la première instance
                // et on l'explose
                var Tab = (Object[])def.GetInstances();
                var ins = (SketchBlockInstance)Tab[0];
                SM.ExplodeSketchBlockInstance(ins);

                // Fermeture de l'esquisse
                SM.AddToDB = false;
                SM.DisplayWhenAdded = true;
                SM.InsertSketch(true);

                //// On supprime la definition du bloc
                //SupprimerDefBloc(mdl, cheminbloc);

                // On renomme l'esquisse
                Esquisse.Name = CONSTANTES.NOM_ESQUISSE_NUMEROTER;

                mdl.eEffacerSelection();

                // On l'active dans toutes les configurations
                Esquisse.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, null);
            }

            if (Esquisse.IsRef())
            {
                // On selectionne l'esquisse, on la cache
                // et on la masque dans le FeatureMgr
                // elle ne sera pas du tout acessible par l'utilisateur
                Esquisse.eSelect();
                mdl.BlankSketch();
                Esquisse.SetUIState((int)swUIStates_e.swIsHiddenInFeatureMgr, true);
                mdl.eEffacerSelection();

                mdl.EditRebuild3();
            }

            return Esquisse;
        }

        private void Nettoyer()
        {
            WindowLog.Ecrire("Nettoyer les modeles :");
            List<ModelDoc2> ListeMdl = new List<ModelDoc2>(MdlBase.pListerComposants().Keys);

            foreach (var mdl in ListeMdl)
            {
                if (mdl.TypeDoc() != eTypeDoc.Piece) continue;


                foreach (Configuration Cf in mdl.eListeConfigs(eTypeConfig.Depliee))
                    Cf.eSupprimerConfigAvecEtatAff(mdl);

                mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                {
                    mdl.ePropSuppr(CONST_PRODUCTION.ID_PIECE);
                    mdl.ePropSuppr(CONST_PRODUCTION.PIECE_ID_DOSSIERS);
                    mdl.ePropSuppr(CONST_PRODUCTION.MAX_INDEXDIM);

                    // On supprime la definition du bloc
                    SupprimerDefBloc(mdl, CheminBlocEsquisseNumeroter());
                }

                foreach (var cfg in mdl.eListeNomConfiguration())
                {
                    mdl.ShowConfiguration2(cfg);
                    mdl.EditRebuild3();
                    var Piece = mdl.ePartDoc();

                    mdl.ePropSuppr(CONST_PRODUCTION.ID_CONFIG, cfg);
                    mdl.ePropSuppr(CONST_PRODUCTION.CONFIG_ID_DOSSIERS, cfg);

                    foreach (var f in Piece.eListeDesFonctionsDePiecesSoudees())
                    {
                        CustomPropertyManager PM = f.CustomPropertyManager;
                        PM.Delete2(CONSTANTES.REF_DOSSIER);
                        PM.Delete2(CONSTANTES.DESC_DOSSIER);
                        PM.Delete2(CONSTANTES.NOM_DOSSIER);

                        String ValParDefaut = mdl.Extension.GetUserPreferenceString((int)swUserPreferenceStringValue_e.swSheetMetalDescription, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
                        PM.ePropAdd(CONSTANTES.DESC_DOSSIER, ValParDefaut);
                    }
                }

                mdl.eFermerSiDifferent(MdlBase);
            }

            int errors = 0;
            int warnings = 0;
            MdlBase.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent + (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, ref errors, ref warnings);

            WindowLog.Ecrire("- Nettoyage terminé");
        }
    }
}
