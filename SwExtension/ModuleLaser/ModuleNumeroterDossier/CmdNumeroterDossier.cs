using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ModuleLaser
{
    namespace ModuleNumeroterDossier
    {
        public class CmdNumeroterDossier : Cmd
        {
            public ModelDoc2 MdlBase = null;
            public Boolean CombinerCorpsIdentiques = false;

            String cheminbloc = "E:\\Mes documents\\SolidWorks\\2018\\Blocs\\Macro\\REPERAGE_DOSSIER.sldblk";
            String NomEsquisse = "REPERAGE_DOSSIER";

            private int indice = 1;

            private int GenRepereDossier
            {
                get
                {
                    return indice++;
                }
            }

            protected override void Command()
            {
                try
                {
                    var ListeCorps = new List<Corps>();

                    var lst = MdlBase.ListerComposants(false, eTypeCorps.Tole | eTypeCorps.Barre);

                    // Liste des dossiers déjà traité
                    HashSet<String> DossierTraite = new HashSet<String>();
                    // Liste des index déjà attribué pour les dimensions
                    Dictionary<String, int> IndexModele = new Dictionary<string, int>();

                    foreach (var mdl in lst.Keys)
                    {
                        mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                        EsquisseRepere(mdl);

                        foreach (var cfg in lst[mdl].Keys)
                        {
                            mdl.ShowConfiguration2(cfg);
                            mdl.EditRebuild3();

                            var piece = mdl.ePartDoc();
                            var NbConfig = lst[mdl][cfg];
                            foreach (var fDossier in piece.eListeDesFonctionsDePiecesSoudees())
                            {
                                BodyFolder Dossier = fDossier.GetSpecificFeature2();
                                eTypeCorps TypeCorps = Dossier.eTypeDeDossier();

                                if (Dossier.IsNull() || Dossier.eNbCorps() == 0 || !(TypeCorps == eTypeCorps.Barre || TypeCorps == eTypeCorps.Tole))
                                    continue;

                                String NomParam = "";

                                var clef = mdl.GetPathName() + "___" + fDossier.GetID();
                                if (!DossierTraite.Contains(clef))
                                {
                                    DossierTraite.Add(clef);

                                    if (!IndexModele.ContainsKey(mdl.GetPathName()))
                                        IndexModele.Add(mdl.GetPathName(), 1);

                                    CustomPropertyManager PM = fDossier.CustomPropertyManager;
                                    NomParam = String.Format("D{0}@{1}", IndexModele[mdl.GetPathName()]++, NomEsquisse);
                                    var propVal = String.Format("P\"{0}@{1}\"", NomParam, mdl.eNomAvecExt());
                                    PM.ePropAdd(CONSTANTES.REF_DOSSIER, propVal);
                                }
                                else
                                {
                                    CustomPropertyManager PM = fDossier.CustomPropertyManager;
                                    String val, result = ""; Boolean wasResolved, link;
                                    PM.Get6(CONSTANTES.REF_DOSSIER, false, out val, out result, out wasResolved, out link);
                                    NomParam = val.Replace("P\"", "").Replace("@" + mdl.eNomAvecExt() + "\"", "");
                                }

                                Dimension param = mdl.Parameter(NomParam);

                                var SwCorps = Dossier.ePremierCorps();
                                if (SwCorps.IsNull()) continue;

                                var MateriauCorps = SwCorps.eGetMateriauCorpsOuPiece(piece, cfg);

                                Boolean Ajoute = false;

                                var nbCorps = Dossier.eNbCorps() * NbConfig;

                                if (CombinerCorpsIdentiques)
                                {
                                    foreach (var CorpsTest in ListeCorps)
                                    {
                                        if ((MateriauCorps != CorpsTest.Materiau) || (TypeCorps != CorpsTest.TypeCorps)) continue;

                                        if (SwCorps.eEstSemblable(CorpsTest.SwCorps))
                                        {
                                            CorpsTest.Nb += nbCorps;

                                            var errors = param.SetSystemValue3(CorpsTest.Repere * 0.001, (int)swSetValueInConfiguration_e.swSetValue_InSpecificConfigurations, cfg);

                                            CorpsTest.AjouterModele(mdl, cfg, fDossier.GetID());
                                            Ajoute = true;
                                            break;
                                        }
                                    }
                                }

                                if (Ajoute == false)
                                {
                                    var rep = GenRepereDossier;
                                    var errors = param.SetSystemValue3(rep * 0.001, (int)swSetValueInConfiguration_e.swSetValue_InSpecificConfigurations, cfg);

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
                    WindowLog.EcrireF("Nb de corps unique : {0}", ListeCorps.Count);
                    int nbtt = 0;
                    foreach (var corps in ListeCorps)
                    {
                        nbtt += corps.Nb;
                        WindowLog.EcrireF("{0} : {1}", corps.Repere, corps.Nb);
                    }

                    WindowLog.EcrireF("Nb total de corps : {0}", nbtt);

                }
                catch (Exception e) { this.LogMethode(new Object[] { e }); }
            }

            private Feature EsquisseRepere(ModelDoc2 mdl)
            {
                Feature Esquisse = mdl.eChercherFonction(fc => { return fc.Name == NomEsquisse; });

                if (Esquisse.IsNull())
                {
                    Feature Plan = mdl.eListeFonctions(fc => { return fc.GetTypeName2() == FeatureType.swTnRefPlane; })[1];

                    Plan.eSelect();

                    var SM = mdl.SketchManager;

                    SM.InsertSketch(true);
                    SM.AddToDB = false;
                    SM.DisplayWhenAdded = true;

                    Esquisse = mdl.Extension.GetLastFeatureAdded();

                    MathUtility Mu = App.Sw.GetMathUtility();
                    MathPoint mp = Mu.CreatePoint(new double[] { 0, 0, 0 });

                    var def = SM.MakeSketchBlockFromFile(mp, cheminbloc, false, 1, 0);

                    if (def.IsNull())
                    {
                        var TabDef = (Object[])SM.GetSketchBlockDefinitions();
                        foreach (SketchBlockDefinition blocdef in TabDef)
                        {
                            if (blocdef.FileName == cheminbloc)
                            {
                                def = blocdef;
                                break;
                            }
                        }
                    }

                    var Tab = (Object[])def.GetInstances();
                    var ins = (SketchBlockInstance)Tab[0];
                    SM.ExplodeSketchBlockInstance(ins);

                    SM.AddToDB = false;
                    SM.DisplayWhenAdded = true;
                    SM.InsertSketch(true);

                    Esquisse.Name = NomEsquisse;
                    
                    mdl.eEffacerSelection();

                    Esquisse.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, null);
                }

                Esquisse.eSelect();
                mdl.BlankSketch();
                Esquisse.SetUIState((int)swUIStates_e.swIsHiddenInFeatureMgr, true);
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
        }
    }
}


