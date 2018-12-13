using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;

namespace ModuleProduction
{
    namespace ModuleRepererDossier
    {
        public class CmdRepererDossier : Cmd
        {
            public ModelDoc2 MdlBase = null;
            public int IndiceCampagne = 0;

            public Boolean CombinerCorpsIdentiques = false;
            public Boolean SupprimerReperes = false;
            public List<Corps> ListeCorpsExistant = new List<Corps>();
            public String FichierNomenclature = "";

            private int _GenRepereDossier = 0;
            private int GenRepereDossier { get { return ++_GenRepereDossier; } }

            // Liste des dossiers déjà traité
            HashSet<String> DossierTraite = new HashSet<String>();
            // Liste des index déjà attribués pour les dimensions
            Dictionary<String, int> IndexDimension = new Dictionary<string, int>();

            /// <summary>
            /// Pour pouvoir obtenir une référence unique pour chaque dossier de corps, identiques
            /// dans l'assemblage, on passe par la création d'une propriété dans chaque dossier.
            /// Cette propriété est liée à une cote dans une esquisse dont la valeur est égale au repère.
            /// Suivant la configuration, la valeur de la cote peut changer et donc le repère du dossier.
            /// C'est le seul moyen pour avoir un lien entre les dossiers et la configuration du modèle.
            /// Les propriétés des dossiers ne sont pas configurables.
            /// </summary>
            protected override void Command()
            {
                try
                {
                    if (SupprimerReperes && (ListeCorpsExistant.Count == 0))
                        NettoyerModele();

                    var ListeCorps = new List<Corps>();

                    var lst = MdlBase.ListerComposants(false);

                    foreach (var mdl in lst.Keys)
                    {
                        mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                        EsquisseRepere(mdl);

                        foreach (var cfg in lst[mdl].Keys)
                        {
                            mdl.ShowConfiguration2(cfg);
                            mdl.EditRebuild3();
                            WindowLog.SautDeLigne();
                            WindowLog.EcrireF("{0} \"{1}\"", mdl.eNomSansExt(), cfg);

                            var piece = mdl.ePartDoc();
                            var NbConfig = lst[mdl][cfg];
                            var ListeDossier = piece.eListeDesFonctionsDePiecesSoudees(
                                swD =>
                                {
                                    BodyFolder Dossier = swD.GetSpecificFeature2();

                                    // Si le dossier est la racine d'un sous-ensemble soudé, il n'y a rien dedans
                                    if (Dossier.IsRef() && Dossier.eNbCorps() > 0 &&
                                        (eTypeCorps.Barre | eTypeCorps.Tole).HasFlag(Dossier.eTypeDeDossier()))
                                        return true;

                                    return false;
                                }
                                );

                            foreach (var fDossier in ListeDossier)
                            {
                                BodyFolder Dossier = fDossier.GetSpecificFeature2();
                                CustomPropertyManager PM = fDossier.CustomPropertyManager;

                                // Si MajDossier et PropExiste, ne pas mettre à jour
                                Boolean Maj = !PM.ePropExiste(CONSTANTES.REF_DOSSIER);

                                String NomParam = "";

                                WindowLog.EcrireF("     {0}", fDossier.Name);

                                // On recherche si le dossier à déjà été traité.
                                //      Si non, on ajoute le dossier à la liste
                                //          On met à jour la liste des index des dimensions :
                                //              On ajoute le nom du modele et on itiniatile l'index à 1
                                //              Ou, puisque c'est un nouveau dossier, on ajoute un à l'index existant.
                                //          On créer le nom du paramètre
                                //          On ajoute la propriété au dossier selon le modèle suivant :
                                //              P"D1@REPERAGE_DOSSIER@Nom_de_la_piece.SLDPRT"
                                //      Si oui, on récupère le nom du paramètre
                                var clef = HashDossier(mdl, fDossier);

                                if (!DossierTraite.Contains(clef))
                                {
                                    DossierTraite.Add(clef);

                                    IndexDimension.AddIfNotExistOrPlus(mdl.GetPathName());

                                    NomParam = String.Format("D{0}@{1}", IndexDimension[mdl.GetPathName()]++, CONSTANTES.NOM_ESQUISSE_NUMEROTER);
                                    var propVal = String.Format("{0}\"{1}@{2}\"", CONSTANTES.PREFIXE_REF_DOSSIER, NomParam, mdl.eNomAvecExt());
                                    var r = PM.ePropAdd(CONSTANTES.REF_DOSSIER, propVal);

                                    if (r > 0)
                                        WindowLog.EcrireF("{0}-{1}-{2} : Pas de propriété ajoutée {3}", mdl.eNomSansExt(), cfg, fDossier.Name, (swCustomInfoAddResult_e)r);

                                    PM.ePropAdd(CONSTANTES.DESC_DOSSIER, propVal);
                                }
                                else
                                {
                                    String val, result = ""; Boolean wasResolved, link;
                                    var r = PM.Get6(CONSTANTES.REF_DOSSIER, false, out val, out result, out wasResolved, out link);

                                    if (r == 1)
                                        WindowLog.EcrireF("{0}-{1}-{2} : Pas de propriété {3}", mdl.eNomSansExt(), cfg, fDossier.Name, (swCustomInfoGetResult_e)r);

                                    NomParam = ExtractNomParam(val);
                                }

                                {
                                    String val, result = ""; Boolean wasResolved, link;
                                    var r = PM.Get6(CONSTANTES.REF_DOSSIER, false, out val, out result, out wasResolved, out link);

                                    PM.ePropAdd(CONSTANTES.DESC_DOSSIER, val);
                                }

                                // On ajoute la propriété NomDossier
                                // permettant de récupérer le nom du dossier dans la mise en plan
                                if (!PM.ePropExiste(CONSTANTES.NOM_DOSSIER))
                                {
                                    var propVal = String.Format("\"SW-CutListItemName@@@{0}@{1}\"", fDossier.Name, mdl.eNomAvecExt());
                                    PM.ePropAdd(CONSTANTES.NOM_DOSSIER, propVal);
                                }

                                var SwCorps = Dossier.ePremierCorps();

                                Boolean Ajoute = false;

                                var MateriauCorps = SwCorps.eGetMateriauCorpsOuPiece(piece, cfg);
                                var nbCorps = Dossier.eNbCorps() * NbConfig;
                                Dimension param = mdl.Parameter(NomParam);

                                eTypeCorps TypeCorps = Dossier.eTypeDeDossier();

                                if (CombinerCorpsIdentiques)
                                {
                                    // On recherche s'il existe des corps identiques
                                    // Si oui, on applique le même repère au parametre
                                    foreach (var CorpsTest in ListeCorps)
                                    {
                                        if ((MateriauCorps != CorpsTest.Materiau) || (TypeCorps != CorpsTest.TypeCorps)) continue;

                                        if (SwCorps.eEstSemblable(CorpsTest.SwCorps))
                                        {
                                            CorpsTest.Nb += nbCorps;

                                            if (Maj)
                                            {
                                                var errors = param.SetSystemValue3(CorpsTest.Repere * 0.001, (int)swSetValueInConfiguration_e.swSetValue_InSpecificConfigurations, cfg);
                                                if (errors > 0)
                                                    WindowLog.EcrireF(" Erreur de mise à jour {0}", (swSetValueReturnStatus_e)errors);
                                            }

                                            //CorpsTest.AjouterModele(mdl, cfg, fDossier.GetID());
                                            Ajoute = true;
                                            break;
                                        }
                                    }
                                }

                                if ((Ajoute == false) && Maj)
                                {
                                    var rep = GenRepereDossier;
                                    var errors = param.SetSystemValue3(rep * 0.001, (int)swSetValueInConfiguration_e.swSetValue_InSpecificConfigurations, cfg);

                                    if (errors > 0)
                                        WindowLog.EcrireF(" Erreur de mise à jour {0}", (swSetValueReturnStatus_e)errors);

                                    var corps = new Corps(SwCorps, TypeCorps, MateriauCorps);
                                    corps.Campagne = IndiceCampagne;
                                    corps.Nb = nbCorps;
                                    corps.Repere = rep;
                                    if (corps.TypeCorps == eTypeCorps.Tole)
                                        corps.Dimension = SwCorps.eEpaisseurCorpsOuDossier(Dossier).ToString();
                                    else
                                        corps.Dimension = Dossier.eProfilDossier();
                                    //corps.AjouterModele(mdl, cfg, fDossier.GetID());
                                    ListeCorps.Add(corps);
                                }
                            }
                        }

                        if (mdl.GetPathName() != MdlBase.GetPathName())
                            App.Sw.CloseDoc(mdl.GetPathName());
                    }

                    MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);


                    int Errors = 0, Warnings = 0;
                    MdlBase.Save3((int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced + (int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref Errors, ref Warnings);
                    MdlBase.EditRebuild3();

                    // Petit récap
                    WindowLog.SautDeLigne();
                    WindowLog.EcrireF("Nb de corps unique : {0}", ListeCorps.Count);

                    int nbtt = 0;

                    using (var sw = new StreamWriter(FichierNomenclature, true))
                    {
                        foreach (var corps in ListeCorps)
                        {
                            nbtt += corps.Nb;
                            WindowLog.EcrireF("P{0} ×{1}", corps.Repere, corps.Nb);
                            sw.WriteLine(String.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}", corps.Campagne, corps.Repere, corps.Nb, corps.TypeCorps, corps.Dimension, corps.Materiau));
                        }
                    }

                    WindowLog.EcrireF("Nb total de corps : {0}", nbtt);
                }
                catch (Exception e)
                {
                    this.LogErreur(new Object[] { e });
                }
            }

            private String HashDossier(ModelDoc2 mdl, Feature f)
            {
                return mdl.GetPathName() + "___" + f.GetID();
            }

            private String ExtractNomParam(String s)
            {
                // val.Replace(CONSTANTES.PREFIXE_REF_DOSSIER + "\"", "").Replace("@" + mdl.eNomAvecExt() + "\"", "");
                s = s.Replace(CONSTANTES.PREFIXE_REF_DOSSIER + "\"", "").Replace("\"", "");
                var t = s.Split('@');
                if (t.Length > 2)
                    return String.Format("{0}@{1}", t[0], t[1]);

                this.LogErreur(new Object[] { "Pas de parametre dans la reference dossier" });
                return "";
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

            private void NettoyerModele()
            {
                WindowLog.Ecrire("Nettoyer les modeles");
                List<ModelDoc2> ListeComposants = new List<ModelDoc2>();
                if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                    ListeComposants.Add(MdlBase);
                else
                {
                    var l = MdlBase.eListeComposants();
                    foreach (var cp in l)
                        ListeComposants.Add(cp.eModelDoc2());
                }

                Predicate<Feature> Test = delegate (Feature f)
                {
                    BodyFolder dossier = f.GetSpecificFeature2();
                    if (dossier.IsRef() && dossier.eNbCorps() > 0)
                        return true;

                    return false;
                };

                foreach (var mdl in ListeComposants)
                {
                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                    {
                        CustomPropertyManager PM = mdl.Extension.CustomPropertyManager[""];
                        PM.Delete2(CONSTANTES.ID_Piece);

                        // On supprime la definition du bloc
                        SupprimerDefBloc(mdl, CheminBlocEsquisseNumeroter());
                    }

                    foreach (var cfg in mdl.eListeNomConfiguration())
                    {
                        mdl.ShowConfiguration2(cfg);
                        mdl.EditRebuild3();
                        var Piece = mdl.ePartDoc();

                        {
                            CustomPropertyManager PM = mdl.Extension.CustomPropertyManager[cfg];
                            PM.Delete2(CONSTANTES.ID_Config);
                        }

                        foreach (var f in Piece.eListeDesFonctionsDePiecesSoudees(Test))
                        {
                            CustomPropertyManager PM = f.CustomPropertyManager;
                            PM.Delete2(CONSTANTES.REF_DOSSIER);
                            PM.Delete2(CONSTANTES.DESC_DOSSIER);
                            PM.Delete2(CONSTANTES.NOM_DOSSIER);
                        }
                    }

                    if (mdl.GetPathName() != MdlBase.GetPathName())
                        App.Sw.CloseDoc(mdl.GetPathName());
                }

                int errors = 0;
                int warnings = 0;
                MdlBase.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent + (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, ref errors, ref warnings);

                WindowLog.Ecrire("Nettoyage terminé");
            }
        }
    }
}


