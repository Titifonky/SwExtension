using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;

namespace ModuleLaser
{
    namespace ModuleExportBarre
    {
        public class CmdExportBarre : Cmd
        {
            public ModelDoc2 MdlBase = null;
            public List<String> ListeMateriaux = new List<String>();
            public int Quantite = 1;
            public Boolean CreerPdf3D = false;
            public eTypeFichierExport TypeExport = eTypeFichierExport.ParasolidBinary;
            public Boolean PrendreEnCompteTole = false;
            public Boolean ComposantsExterne = false;
            public String RefFichier = "";

            public Boolean ReinitialiserNoDossier = false;
            public Boolean MajListePiecesSoudees = false;
            public String ForcerMateriau = null;

            private String DossierExport = "";
            private String DossierExportPDF = "";
            private String Indice = "";
            private HashSet<String> HashMateriaux;
            private Dictionary<String, int> DicQte = new Dictionary<string, int>();
            private Dictionary<String, List<String>> DicConfig = new Dictionary<String, List<String>>();
            private List<Component2> ListeCp = new List<Component2>();

            private InfosBarres Nomenclature = new InfosBarres();

            protected override void Command()
            {
                CreerDossierDVP();

                WindowLog.Ecrire(String.Format("Dossier :\r\n{0}", new DirectoryInfo(DossierExport).Name));

                try
                {
                    HashMateriaux = new HashSet<string>(ListeMateriaux);

                    if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                    {
                        Component2 CpRacine = MdlBase.eComposantRacine();
                        ListeCp.Add(CpRacine);

                        if (!CpRacine.eNomConfiguration().eEstConfigPliee())
                        {
                            WindowLog.Ecrire("Pas de configuration valide," +
                                                "\r\n le nom de la config doit être composée exclusivement de chiffres");
                            return;
                        }

                        List<String> ListeConfig = new List<String>() { CpRacine.eNomConfiguration() };

                        DicConfig.Add(CpRacine.eKeySansConfig(), ListeConfig);

                        if (ReinitialiserNoDossier)
                        {
                            var nomConfigBase = MdlBase.eNomConfigActive();
                            foreach (var cfg in ListeConfig)
                            {
                                MdlBase.ShowConfiguration2(cfg);
                                MdlBase.eComposantRacine().eEffacerNoDossier();
                            }
                            MdlBase.ShowConfiguration2(nomConfigBase);
                        }

                        DicQte.Add(CpRacine.eKeyAvecConfig());
                    }

                    eTypeCorps Filtre = PrendreEnCompteTole ? eTypeCorps.Barre | eTypeCorps.Tole : eTypeCorps.Barre;



                    // Si c'est un assemblage, on liste les composants
                    if (MdlBase.TypeDoc() == eTypeDoc.Assemblage)
                        ListeCp = MdlBase.eRecListeComposant(
                        c =>
                        {
                            if ((ComposantsExterne || c.eEstDansLeDossier(MdlBase)) && !c.IsHidden(true) && !c.ExcludeFromBOM && (c.TypeDoc() == eTypeDoc.Piece))
                            {
                                if (!c.eNomConfiguration().eEstConfigPliee() || DicQte.Plus(c.eKeyAvecConfig()))
                                    return false;

                                if (ReinitialiserNoDossier)
                                    c.eEffacerNoDossier();

                                var LstDossier = c.eListeDesDossiersDePiecesSoudees();
                                foreach (var dossier in LstDossier)
                                {
                                    if (!dossier.eEstExclu() && Filtre.HasFlag(dossier.eTypeDeDossier()))
                                    {
                                        String Materiau = dossier.eGetMateriau();

                                        if (!HashMateriaux.Contains(Materiau))
                                            continue;

                                        DicQte.Add(c.eKeyAvecConfig());

                                        if (DicConfig.ContainsKey(c.eKeySansConfig()))
                                        {
                                            List<String> l = DicConfig[c.eKeySansConfig()];
                                            if (l.Contains(c.eNomConfiguration()))
                                                return false;
                                            else
                                            {
                                                l.Add(c.eNomConfiguration());
                                                return true;
                                            }
                                        }
                                        else
                                        {
                                            DicConfig.Add(c.eKeySansConfig(), new List<string>() { c.eNomConfiguration() });
                                            return true;
                                        }
                                    }
                                }

                                //foreach (Body2 corps in c.eListeCorps())
                                //{
                                //    if (Filtre.HasFlag(corps.eTypeDeCorps()))
                                //    {
                                //        String Materiau = corps.eGetMateriauCorpsOuComp(c);

                                //        if (!HashMateriaux.Contains(Materiau))
                                //            continue;

                                //        DicQte.Add(c.eKeyAvecConfig());

                                //        if (DicConfig.ContainsKey(c.eKeySansConfig()))
                                //        {
                                //            if(DicConfig[c.eKeySansConfig()].AddIfNotExist(c.eNomConfiguration()) && ReinitialiserNoDossier)
                                //                c.eEffacerNoDossier();

                                //            return false;
                                //        }

                                //        DicConfig.Add(c.eKeySansConfig(), new List<string>() { c.eNomConfiguration() });
                                //        return true;
                                //    }
                                //}
                            }

                            return false;
                        },
                        null,
                        true
                    );

                    Nomenclature.TitreColonnes("Barre ref.", "Materiau", "Profil", "Lg", "Nb");

                    // On multiplie les quantites
                    DicQte.Multiplier(Quantite);

                    for (int noCp = 0; noCp < ListeCp.Count; noCp++)
                    {
                        var Cp = ListeCp[noCp];
                        ModelDoc2 mdl = Cp.eModelDoc2();
                        mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                        var ListeNomConfigs = DicConfig[Cp.eKeySansConfig()];
                        ListeNomConfigs.Sort(new WindowsStringComparer());

                        if (ReinitialiserNoDossier)
                            mdl.ePartDoc().eReinitialiserNoDossierMax();

                        WindowLog.SautDeLigne();
                        WindowLog.EcrireF("[{1}/{2}] {0}", Cp.eNomSansExt(), noCp + 1, ListeCp.Count);

                        for (int noCfg = 0; noCfg < ListeNomConfigs.Count; noCfg++)
                        {
                            var NomConfigPliee = ListeNomConfigs[noCfg];
                            int QuantiteCfg = DicQte[Cp.eKeyAvecConfig(NomConfigPliee)];
                            WindowLog.SautDeLigne();
                            WindowLog.EcrireF("  [{2}/{3}] Config : \"{0}\" -> ×{1}", NomConfigPliee, QuantiteCfg, noCfg + 1, ListeNomConfigs.Count);
                            mdl.ShowConfiguration2(NomConfigPliee);
                            mdl.EditRebuild3();
                            PartDoc Piece = mdl.ePartDoc();

                            ListPID<Feature> ListeDossier = Piece.eListePIDdesFonctionsDePiecesSoudees(null);

                            for (int noD = 0; noD < ListeDossier.Count; noD++)
                            {
                                Feature f = ListeDossier[noD];
                                BodyFolder dossier = f.GetSpecificFeature2();

                                if (dossier.eEstExclu() || dossier.IsNull() || (dossier.GetBodyCount() == 0)) continue;

                                WindowLog.SautDeLigne();
                                WindowLog.EcrireF("    - [{1}/{2}] Dossier : \"{0}\"", f.Name, noD + 1, ListeDossier.Count);

                                Body2 Barre = dossier.ePremierCorps();

                                String Profil = dossier.eProp(CONSTANTES.PROFIL_NOM);
                                String Longueur = dossier.eProp(CONSTANTES.PROFIL_LONGUEUR);

                                if (String.IsNullOrWhiteSpace(Profil) || String.IsNullOrWhiteSpace(Longueur))
                                {
                                    WindowLog.Ecrire("      Pas de barres");
                                    continue;
                                }

                                String Materiau = Barre.eGetMateriauCorpsOuPiece(Piece, NomConfigPliee);

                                if (!HashMateriaux.Contains(Materiau))
                                {
                                    WindowLog.Ecrire("      Pas de barres");
                                    continue;
                                }

                                Materiau = ForcerMateriau.IsRefAndNotEmpty(Materiau);

                                String noDossier = dossier.eProp(CONSTANTES.NO_DOSSIER);

                                if (noDossier.IsNull() || String.IsNullOrWhiteSpace(noDossier))
                                    noDossier = Piece.eNumeroterDossier(MajListePiecesSoudees)[dossier.eNom()].ToString();

                                int QuantiteBarre = QuantiteCfg * dossier.GetBodyCount();

                                String RefBarre = ConstruireRefBarre(mdl, NomConfigPliee, noDossier);
                                String NomFichierBarre = ConstruireNomFichierBarre(RefBarre, QuantiteBarre);

                                WindowLog.EcrireF("    Profil {0}  Materiau {1}", Profil, Materiau);



                                Nomenclature.AjouterLigne(RefBarre, Materiau, Profil, Math.Round(Longueur.eToDouble()).ToString(), "× " + QuantiteBarre.ToString());

                                //mdl.ViewZoomtofit2();
                                //mdl.ShowNamedView2("*Isométrique", 7);

                                ModelDoc2 mdlBarre = Barre.eEnregistrerSous(Piece, DossierExport, NomFichierBarre, TypeExport);

                                if (CreerPdf3D)
                                {
                                    String CheminPDF = Path.Combine(DossierExportPDF, NomFichierBarre + eTypeFichierExport.PDF.GetEnumInfo<ExtFichier>());
                                    mdlBarre.SauverEnPdf3D(CheminPDF);
                                }

                                App.Sw.CloseDoc(mdlBarre.GetPathName());
                            }
                        }

                        if (Cp.GetPathName() != MdlBase.GetPathName())
                            App.Sw.CloseDoc(mdl.GetPathName());
                    }

                    WindowLog.SautDeLigne();
                    WindowLog.Ecrire(Nomenclature.ListeLignes());

                    StreamWriter s = new StreamWriter(Path.Combine(DossierExport, "Nomenclature.txt"));
                    s.Write(Nomenclature.GenererTableau());
                    s.Close();

                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            public String ConstruireRefBarre(ModelDoc2 mdl, String configPliee, String noDossier)
            {
                return String.Format("{0} - {1}-{2}-{3}", RefFichier, mdl.eNomSansExt(), configPliee, noDossier);
            }

            public String ConstruireNomFichierBarre(String reBarre, int quantite)
            {
                return String.Format("{0} (×{1}) - {2}", reBarre, quantite, Indice);
            }

            private void CreerDossierDVP()
            {
                String NomBase = RefFichier + " - " + CONSTANTES.DOSSIER_BARRE + "_" + TypeExport.GetEnumInfo<ExtFichier>().Replace(".", "").ToUpperInvariant();

                DirectoryInfo D = new DirectoryInfo(MdlBase.eDossier());
                List<String> ListeD = new List<string>();

                foreach (var d in D.GetDirectories())
                {
                    if (d.Name.ToUpperInvariant().StartsWith(NomBase.ToUpperInvariant()))
                    {
                        ListeD.Add(d.Name);
                    }
                }

                ListeD.Sort(new WindowsStringComparer(ListSortDirection.Ascending));

                Indice = ChercherIndice(ListeD);

                DossierExport = Path.Combine(MdlBase.eDossier(), NomBase + " - " + Indice);

                if (!Directory.Exists(DossierExport))
                    Directory.CreateDirectory(DossierExport);

                if (CreerPdf3D)
                {
                    DossierExportPDF = Path.Combine(DossierExport, "PDF");

                    if (!Directory.Exists(DossierExportPDF))
                        Directory.CreateDirectory(DossierExportPDF);
                }

            }

            private readonly String ChaineIndice = "ZYXWVUTSRQPONMLKJIHGFEDCBA";

            private String ChercherIndice(List<String> liste)
            {
                for (int i = 0; i < ChaineIndice.Length; i++)
                {
                    if (liste.Any(d => { return d.EndsWith(" Ind " + ChaineIndice[i]) ? true : false; }))
                        return "Ind " + ChaineIndice[Math.Max(0, i - 1)];
                }

                return "Ind " + ChaineIndice.Last();
            }

            private class InfosBarres : List<List<String>>
            {
                private List<String> _TitreColonnes = new List<string>();
                private List<int> _DimColonnes = new List<int>();

                public void TitreColonnes(params String[] Valeurs)
                {
                    for (int i = 0; i < Valeurs.Length; i++)
                    {
                        if (i < _DimColonnes.Count)
                            _DimColonnes[i] = Math.Max(_DimColonnes[i], Valeurs[i].Length);
                        else
                            _DimColonnes.Add(Valeurs[i].Length);
                    }

                    _TitreColonnes = new List<string>(Valeurs);
                }

                public void AjouterLigne(params String[] Valeurs)
                {
                    for (int i = 0; i < Valeurs.Length; i++)
                    {
                        if (i < _DimColonnes.Count)
                            _DimColonnes[i] = Math.Max(_DimColonnes[i], Valeurs[i].Length);
                        else
                            _DimColonnes.Add(Valeurs[i].Length);
                    }

                    Add(new List<string>(Valeurs));
                }

                public List<String> ListeLignes()
                {
                    List<String> Liste = new List<string>();

                    if (_TitreColonnes.Count != 0)
                    {
                        String formatTitre = "";

                        for (int i = 0; i < _TitreColonnes.Count; i++)
                            formatTitre += "{" + i.ToString() + ",-" + _DimColonnes[i] + "}    ";

                        formatTitre = formatTitre.Trim();

                        Liste.Add(String.Format(formatTitre, _TitreColonnes.ToArray()));
                    }

                    if (Count != 0)
                    {

                        String format = "";

                        for (int i = 0; i < _DimColonnes.Count; i++)
                        {
                            String Sign = "";
                            if (Char.IsLetter(this[0][i].Trim()[0]))
                                Sign = "-";

                            format += "{" + i.ToString() + "," + Sign + _DimColonnes[i] + "}    ";
                        }

                        format = format.Trim();

                        foreach (List<String> ligne in this)
                        {
                            Liste.Add(String.Format(format, ligne.ToArray()));
                        }
                    }

                    return Liste;
                }

                public String GenererTableau()
                {
                    return String.Join("\r\n", ListeLignes());
                }
            }
        }
    }
}


