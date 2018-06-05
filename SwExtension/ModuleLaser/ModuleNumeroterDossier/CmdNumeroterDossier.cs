using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace ModuleLaser
{
    namespace ModuleNumeroterDossier
    {
        public class CmdNumeroterDossier : Cmd
        {
            public ModelDoc2 MdlBase = null;
            public Boolean CombinerCorpsIdentiques = false;
            public Boolean MajDossiers = false;

            private int indice = 0;

            private int GenRepereDossier { get { return ++indice; } }

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
                    if (MajDossiers)
                        IndiceMax();

                    var ListeCorps = new List<Corps>();

                    var lst = MdlBase.ListerComposants(false, eTypeCorps.Tole | eTypeCorps.Barre);

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
                                    if (Dossier.IsRef() && Dossier.eNbCorps() > 0)
                                    {
                                        eTypeCorps TypeCorps = Dossier.eTypeDeDossier();
                                        if (TypeCorps == eTypeCorps.Barre || TypeCorps == eTypeCorps.Tole)
                                            return true;
                                    }

                                    return false;
                                }
                                );

                            foreach (var fDossier in ListeDossier)
                            {
                                BodyFolder Dossier = fDossier.GetSpecificFeature2();
                                CustomPropertyManager PM = fDossier.CustomPropertyManager;

                                // Si MajDossier et PropExiste, ne pas mettre à jour
                                Boolean Maj = !(MajDossiers && PM.ePropExiste(CONSTANTES.REF_DOSSIER));

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

                                            CorpsTest.AjouterModele(mdl, cfg, fDossier.GetID());
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
                                    corps.Nb = nbCorps;
                                    corps.Repere = rep;
                                    corps.AjouterModele(mdl, cfg, fDossier.GetID());
                                    ListeCorps.Add(corps);
                                }
                            }
                        }
                    }

                    MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);


                    int Errors = 0, Warnings = 0;
                    MdlBase.Save3((int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced + (int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref Errors, ref Warnings);
                    MdlBase.EditRebuild3();

                    // Petit récap
                    WindowLog.SautDeLigne();
                    WindowLog.EcrireF("Nb de corps unique : {0}", ListeCorps.Count);

                    int nbtt = 0;

                    foreach (var corps in ListeCorps)
                    {
                        nbtt += corps.Nb;
                        WindowLog.EcrireF("P{0} ×{1}", corps.Repere, corps.Nb);
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

            private Feature EsquisseRepere(ModelDoc2 mdl)
            {
                // On recherche l'esquisse contenant les parametres
                Feature Esquisse = mdl.eChercherFonction(fc => { return fc.Name == CONSTANTES.NOM_ESQUISSE_NUMEROTER; });

                if (Esquisse.IsNull())
                {
                    // On recherche le plan de dessus, le deuxième dans la liste des plans de référence
                    Feature Plan = mdl.eListeFonctions(fc => { return fc.GetTypeName2() == FeatureType.swTnRefPlane; })[1];

                    // Selection du plan et création de l'esquisse
                    Plan.eSelect();
                    var SM = mdl.SketchManager;
                    SM.InsertSketch(true);
                    SM.AddToDB = false;
                    SM.DisplayWhenAdded = true;

                    mdl.eEffacerSelection();

                    // On récupère la fonction de l'esquisse
                    Esquisse = mdl.Extension.GetLastFeatureAdded();

                    // On recherche le chemin du bloc
                    String cheminbloc = "";

                    var CheminDossierBloc = App.Sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsBlocks).Split(new String[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (var chemin in CheminDossierBloc)
                    {
                        var d = new DirectoryInfo(chemin);
                        var r = d.GetFiles(CONSTANTES.NOM_BLOCK_ESQUISSE_NUMEROTER, SearchOption.AllDirectories);
                        if (r.Length > 0)
                        {
                            cheminbloc = r[0].FullName;
                            break;
                        }
                    }

                    if (String.IsNullOrWhiteSpace(cheminbloc))
                        return null;

                    // On insère le bloc
                    MathUtility Mu = App.Sw.GetMathUtility();
                    MathPoint Origine = Mu.CreatePoint(new double[] { 0, 0, 0 });
                    var def = SM.MakeSketchBlockFromFile(Origine, cheminbloc, false, 1, 0);

                    // Si l'insertion a échoué, c'est qu'il existe
                    // déjà une definition dans le modèle
                    // On la recherche
                    if (def.IsNull())
                    {
                        var TabDef = (Object[])SM.GetSketchBlockDefinitions();
                        foreach (SketchBlockDefinition blocdef in TabDef)
                        {
                            WindowLog.Ecrire(blocdef.FileName);
                            if (blocdef.FileName == cheminbloc)
                            {
                                def = blocdef;
                                break;
                            }
                        }

                        // On insère le bloc
                        SM.InsertSketchBlockInstance(def, Origine, 1, 0);
                    }

                    // On récupère la première instance
                    // et on l'explose
                    var Tab = (Object[])def.GetInstances();
                    var ins = (SketchBlockInstance)Tab[0];
                    SM.ExplodeSketchBlockInstance(ins);

                    // Fermeture de l'esquisse
                    SM.AddToDB = false;
                    SM.DisplayWhenAdded = true;
                    SM.InsertSketch(true);

                    // On renomme l'esquisse
                    Esquisse.Name = CONSTANTES.NOM_ESQUISSE_NUMEROTER;

                    mdl.eEffacerSelection();

                    // On l'active dans toutes les configurations
                    Esquisse.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, null);
                }

                // On selectionne l'esquisse, on la cache
                // et on la masque dans le FeatureMgr
                // elle ne sera pas du tout acessible par l'utilisateur
                Esquisse.eSelect();
                mdl.BlankSketch();
                Esquisse.SetUIState((int)swUIStates_e.swIsHiddenInFeatureMgr, true);
                mdl.eEffacerSelection();
                return Esquisse;
            }

            private class Corps
            {
                public Body2 SwCorps;
                public eTypeCorps TypeCorps;
                public String Materiau;
                public int Repere;
                public SortedDictionary<ModelDoc2, SortedDictionary<String, Dictionary<int, int>>> ListeModele = new SortedDictionary<ModelDoc2, SortedDictionary<String, Dictionary<int, int>>>(new CompareModelDoc2());
                public int Nb = 0;

                public Corps(Body2 swCorps, eTypeCorps typeCorps, String materiau)
                {
                    SwCorps = swCorps;
                    TypeCorps = typeCorps;
                    Materiau = materiau;
                }

                public void AjouterModele(ModelDoc2 mdl, String config, int dossier)
                {
                    if (ListeModele.ContainsKey(mdl))
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
                            var lDossier = new Dictionary<int, int>();
                            lDossier.Add(dossier, Repere);
                            lCfg.Add(config, lDossier);
                        }
                    }
                    else
                    {
                        var lDossier = new Dictionary<int, int>();
                        lDossier.Add(dossier, Repere);
                        var lCfg = new SortedDictionary<String, Dictionary<int, int>>(new WindowsStringComparer());
                        lCfg.Add(config, lDossier);
                        ListeModele.Add(mdl, lCfg);
                    }
                }

                public void AjouterModele(Component2 comp, int dossier)
                {
                    AjouterModele(comp.eModelDoc2(), comp.eNomConfiguration(), dossier);
                }
            }

            private void IndiceMax()
            {
                eTypeCorps Filtre = eTypeCorps.Barre | eTypeCorps.Tole;

                var Dic = new HashSet<String>();

                Predicate<Component2> Test = delegate (Component2 comp)
                {
                    if (!comp.IsSuppressed())
                    {
                        var hashComp = comp.eKeyAvecConfig();
                        if (!Dic.Contains(hashComp))
                        {
                            Dic.Add(hashComp);

                            var l = comp.eListeDesFonctionsDePiecesSoudees(
                                f =>
                                {
                                    BodyFolder dossier = f.GetSpecificFeature2();
                                    if (dossier.IsRef() && dossier.eNbCorps() > 0 && Filtre.HasFlag(dossier.eTypeDeDossier()))
                                    {
                                        CustomPropertyManager PM = f.CustomPropertyManager;

                                        String val, result = ""; Boolean wasResolved, link;
                                        var r = PM.Get6(CONSTANTES.REF_DOSSIER, false, out val, out result, out wasResolved, out link);

                                        var RefDossier = result.Replace(CONSTANTES.PREFIXE_REF_DOSSIER, "");
                                        if (RefDossier.eIsInteger())
                                        {
                                            indice = Math.Max(indice, RefDossier.eToInteger());
                                            var hashDossier = HashDossier(comp.eModelDoc2(), f);
                                            DossierTraite.Add(hashDossier);

                                            // On recherche l'index max pour chaque modele
                                            var NomParam = ExtractNomParam(val);
                                            var dim = NomParam.Split('@')[0].CleanStringOfNonDigits().eToInteger();
                                            var hashMdl = comp.GetPathName();
                                            if (IndexDimension.ContainsKey(hashMdl))
                                                IndexDimension[hashMdl] = Math.Max(IndexDimension[hashMdl], dim);
                                            else
                                                IndexDimension[hashMdl] = dim;
                                        }
                                    }

                                    return true;
                                }
                                );
                        }

                    }
                    return false;
                };

                if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                    Test(MdlBase.eComposantRacine());
                else
                    MdlBase.eRecParcourirComposants(Test);

                WindowLog.EcrireF("Indice Max : {0}", indice);
            }
        }
    }
}


