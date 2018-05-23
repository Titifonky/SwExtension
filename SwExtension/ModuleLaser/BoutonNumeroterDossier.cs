using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
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
                var r = "P" + (indice++).ToString();

                //while (mdl.FeatureManager.IsNameUsed((int)swNameType_e.swFeatureName, r))

                return r;
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

                        if (Dossier.IsNull() || Dossier.eNbCorps() == 0 || Dossier.eEstExclu() || !(Dossier.eEstUnDossierDeBarres() || Dossier.eEstUnDossierDeToles()))
                            continue;

                        fDossier.Name = "P" + (indice++).ToString();
                    }
                }
                else
                {
                    var init = 1;
                    mdlBase.eComposantRacine().eRecParcourirComposantBase(
                        comp =>
                        {
                            foreach (var fDossier in comp.eListeDesFonctionsDePiecesSoudees())
                            {
                                if (fDossier.Name.StartsWith("P"))
                                    fDossier.Name = "X" + (init++).ToString();
                            }
                        }
                        );

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
                                        NommerDossier(fDossier, CorpsTest.Repere);

                                        if(fDossier.Name.Trim() != CorpsTest.Repere)
                                            WindowLog.EcrireF("Erreur {0}-{1} : {2}->{3}", comp.Name2, comp.eNomConfiguration(), fDossier.Name, CorpsTest.Repere);

                                        Ajoute = true;
                                        break;
                                    }
                                }

                                if (Ajoute == false)
                                {
                                    var rep = GenRepereDossier;
                                    NommerDossier(fDossier, rep);

                                    if (fDossier.Name.Trim() != rep)
                                        WindowLog.EcrireF("  Erreur {0}-{1} : {2}->{3}", comp.Name2, comp.eNomConfiguration(), fDossier.Name, rep);

                                    var corps = new Corps(SwCorps, MateriauCorps);
                                    corps.Nb = Dossier.eNbCorps();
                                    corps.Repere = rep;
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
            }
            catch (Exception e) { this.LogMethode(new Object[] { e }); }
        }

        // On rajoute des espaces au nom pour remedier au problèmes
        // des corps identiques dans des dossiers différents
        private void NommerDossier(Feature f, String rep)
        {
            int Boucle = 0;
            var rtmp = rep;
            f.Name = rtmp;
            while((f.Name != rtmp) && (Boucle++ < 15))
            {
                rtmp += " ";
                f.Name = rtmp;
            }
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
