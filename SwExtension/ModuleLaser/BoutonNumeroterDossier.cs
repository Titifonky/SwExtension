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
        private static AttributeDef AttDef = null;
        private const String ATTRIBUT_LASER_REF = "LaserRef";
        private const String ATTRIBUT_PARAM_REF = "Ref";

        protected override void Command()
        {
            try
            {
                ModelDoc2 mdlBase = App.ModelDoc2;
                var ListeCorps = new List<Corps>();

                int indice = 1;

                if (mdlBase.TypeDoc() == eTypeDoc.Piece)
                {
                    foreach (var fDossier in mdlBase.ePartDoc().eListeDesFonctionsDePiecesSoudees())
                    {
                        BodyFolder Dossier = fDossier.GetSpecificFeature2();

                        if (Dossier.IsNull() || Dossier.eNbCorps() == 0 || Dossier.eEstExclu() || !(Dossier.eEstUnDossierDeBarres() || Dossier.eEstUnDossierDeToles()))
                            continue;

                        fDossier.Name = "P" + (indice++).ToString();
                    }
                }
                else
                {
                    mdlBase.eComposantRacine().eRecParcourirComposantBase(
                        comp =>
                        {
                            foreach (var fDossier in comp.eListeDesFonctionsDePiecesSoudees())
                            {
                                BodyFolder Dossier = fDossier.GetSpecificFeature2();
                                if (Dossier.IsNull() || Dossier.eNbCorps() == 0 || Dossier.eEstExclu() || !(Dossier.eEstUnDossierDeBarres() || Dossier.eEstUnDossierDeToles()))
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
                                        CorpsTest.AjouterModele(comp);
                                        fDossier.Name = CorpsTest.Repere;
                                        Ajoute = true;
                                        break;
                                    }
                                }

                                if (Ajoute == false)
                                {
                                    fDossier.Name = "P" + (indice++).ToString();
                                    var corps = new Corps(SwCorps, MateriauCorps);
                                    corps.Nb = Dossier.eNbCorps();
                                    corps.Repere = fDossier.Name;
                                    corps.AjouterModele(comp);
                                    ListeCorps.Add(corps);
                                }
                            }
                        }
                        );


                    WindowLog.EcrireF("Nb de corps unique : {0}", ListeCorps.Count);
                    int nbtt = 0;
                    foreach (var t in ListeCorps)
                    {
                        nbtt += t.Nb;
                        WindowLog.EcrireF("{0} : {1}", t.Repere, t.Nb);
                        foreach (var r in t.ListeModele)
                        {
                            var mdl = r.Key;
                            WindowLog.EcrireF(" {0}", mdl.eNomSansExt());
                            foreach (var cfg in r.Value)
                            {
                                WindowLog.EcrireF("    {0}", cfg);
                            }
                        }
                        
                    }

                    WindowLog.EcrireF("Nb total de corps : {0}", nbtt);
                }

                //if (AttDef.IsNull())
                //{
                //    AttDef = App.Sw.DefineAttribute(ATTRIBUT_LASER_REF);
                //    AttDef.AddParameter(ATTRIBUT_PARAM_REF, (int)swParamType_e.swParamTypeString, 0, 0);
                //    AttDef.Register();
                //}

                //var ListeComp = mdlBase.eComposantRacine().eRecListeComposant(c => { return c.TypeDoc() == eTypeDoc.Piece; }, null, true);

                //foreach (var comp in ListeComp)
                //{
                //    var ListefDossier = comp.eListeDesFonctionsDePiecesSoudees();

                //    for (int i = 0; i < ListefDossier.Count; i++)
                //    {
                //        Feature fDossier = ListefDossier[i];
                //        BodyFolder Dossier = fDossier.GetSpecificFeature2();
                //        if (Dossier.IsNull() || Dossier.eNbCorps() == 0 || Dossier.eEstExclu() || !(Dossier.eEstUnDossierDeBarres() || Dossier.eEstUnDossierDeToles()))
                //            continue;

                //        CustomPropertyManager PM = Dossier.GetFeature().CustomPropertyManager;
                //        PM.Delete2(CONSTANTES.NO_DOSSIER);
                //    }
                //}

                //int noDossierMax = 1;

                //HashSet<String> DicDossier = new HashSet<string>();

                //foreach (var comp in ListeComp)
                //{
                //    WindowLog.EcrireF("{0}", comp.Name2);

                //    ModelDoc2 mdl = comp.eModelDoc2();
                //    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                //    var ListefDossier = mdl.ePartDoc().eListeDesFonctionsDePiecesSoudees();

                //    for (int i = 0; i < ListefDossier.Count; i++)
                //    {
                //        Feature fDossier = ListefDossier[i];
                //        BodyFolder Dossier = fDossier.GetSpecificFeature2();
                //        if (Dossier.IsNull() || Dossier.eNbCorps() == 0 || Dossier.eEstExclu() || !(Dossier.eEstUnDossierDeBarres() || Dossier.eEstUnDossierDeToles()))
                //            continue;

                //        CustomPropertyManager PM = Dossier.GetFeature().CustomPropertyManager;
                //        var ClefDossier = comp.eNomSansExt() + "-" + fDossier.Name;

                //        // Si la propriete existe, on récupère la valeur
                //        if (!DicDossier.Contains(ClefDossier))
                //        {
                //            WindowLog.EcrireF("  {0} -> {1}", fDossier.Name, noDossierMax);
                //            PM.ePropAdd(CONSTANTES.NO_DOSSIER, noDossierMax++);
                //            DicDossier.Add(ClefDossier);
                //        }  
                //    }

                //    if (comp.GetPathName() != mdlBase.GetPathName())
                //        App.Sw.CloseDoc(mdl.GetPathName());
                //}


            }
            catch (Exception e) { this.LogMethode(new Object[] { e }); }
        }

        private class Corps
        {
            public Body2 SwCorps;
            public String Materiau;
            public String Repere;
            public SortedDictionary<ModelDoc2, SortedSet<String>> ListeModele = new SortedDictionary<ModelDoc2, SortedSet<string>>(new CompareModelDoc2());
            public int Nb = 0;

            public Corps(Body2 swCorps, String materiau)
            {
                SwCorps = swCorps;
                Materiau = materiau;
            }

            public void AjouterModele(ModelDoc2 mdl, String config)
            {
                if(ListeModele.ContainsKey(mdl))
                {
                    var l = ListeModele[mdl];
                    if (!l.Contains(config))
                        l.Add(config);
                }
                else
                {
                    var l = new SortedSet<String>(new WindowsStringComparer());
                    l.Add(config);
                    ListeModele.Add(mdl, l);
                }
            }

            public void AjouterModele(Component2 comp)
            {
                AjouterModele(comp.eModelDoc2(), comp.eNomConfiguration());
            }
        }
    }
}
