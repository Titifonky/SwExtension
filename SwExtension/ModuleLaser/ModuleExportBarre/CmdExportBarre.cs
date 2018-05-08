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

                    Nomenclature.TitreColonnes("Barre ref.", "Materiau", "Profil", "Lg", "Nb", "Usinage Ext 1", "Usinage Ext 2", "Détail des Usinage interne");

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
                                    WindowLog.Ecrire("      Materiau exclu");
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

                                var analyse = new AnalyseBarre(Barre);

                                Dictionary<String, Double> Dic = new Dictionary<string, double>();

                                foreach (var u in analyse.ListeFaceUsinageSection)
                                {
                                    String nom = u.ListeFaceDecoupe.Count + " face - Lg " + Math.Round(u.LgUsinage * 1000, 1);
                                    if(Dic.ContainsKey(nom))
                                        Dic[nom] += 1;
                                    else
                                        Dic.Add(nom, 1);
                                }

                                String[] Tab = new String[1 + 7];
                                int i = 0;
                                Tab[i++] = RefBarre; Tab[i++] = Materiau; Tab[i++] = Profil;
                                Tab[i++] = Math.Round(Longueur.eToDouble()).ToString();
                                Tab[i++] = "× " + QuantiteBarre.ToString();
                                Tab[i++] = Math.Round(analyse.ListeFaceUsinageExtremite[0].LgUsinage * 1000, 1).ToString();
                                Tab[i++] = Math.Round(analyse.ListeFaceUsinageExtremite[1].LgUsinage * 1000, 1).ToString();
                                Tab[i] = "";
                                foreach (var nom in Dic.Keys)
                                    Tab[i] += Dic[nom] + "x [" + nom + "]   ";
                                WindowLog.Ecrire(Tab[i]);
                                Nomenclature.AjouterLigne(Tab[0], Tab[1], Tab[2], Tab[3], Tab[4], Tab[5], Tab[6], Tab[7]);

                                //mdl.ViewZoomtofit2();
                                //mdl.ShowNamedView2("*Isométrique", 7);

                                //ModelDoc2 mdlBarre = Barre.eEnregistrerSous(Piece, DossierExport, NomFichierBarre, TypeExport);

                                //if (CreerPdf3D)
                                //{
                                //    String CheminPDF = Path.Combine(DossierExportPDF, NomFichierBarre + eTypeFichierExport.PDF.GetEnumInfo<ExtFichier>());
                                //    mdlBarre.SauverEnPdf3D(CheminPDF);
                                //}

                                //App.Sw.CloseDoc(mdlBarre.GetPathName());
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

            // ========================================================================================

            private class AnalyseBarre
            {
                public Body2 Corps = null;

                public Plan PlanSection;
                public Point ExtremPoint1;
                public Point ExtremPoint2;
                public ListFaceGeom FaceSectionExt = null;
                public List<ListFaceGeom> ListeFaceSectionInt = null;
                public List<ListFaceUsinage> ListeFaceUsinageExtremite = new List<ListFaceUsinage>();
                public List<ListFaceUsinage> ListeFaceUsinageSection = new List<ListFaceUsinage>();

                private List<ListFaceGeom> ListeFaceExtremite = null;

                public AnalyseBarre(Body2 corps)
                {
                    Corps = corps;

                    AnalyserFaces();

                    AnalyserPercages();
                }

                #region ANALYSE DES USINAGES

                private void AnalyserPercages()
                {
                    // On recupère les faces issues des fonctions modifiant le corps
                    List<Face2> ListeFaceSection = new List<Face2>();
                    foreach (var Fonc in Corps.eListeFonctions(null))
                    {
                        if (Fonc.GetTypeName2() != "WeldMemberFeat")
                        {
                            foreach (var f in Fonc.eListeDesFaces())
                            {
                                Body2 cF = f.GetBody();
                                if (cF.eIsSame(Corps))
                                    ListeFaceSection.AddIfNotExist(f);
                            }
                        }
                    }

                    // Recherche des perçages
                    // On ajoute les faces d'extrémité
                    List<Face2> ListeFaceExt = new List<Face2>();
                    foreach (var fl in ListeFaceExtremite)
                        ListeFaceExt.AddRange(fl.ListeFaceSw());

                    // On ajoute les faces d'extrémité
                    ListeFaceSection.AddRange(ListeFaceExt);

                    // On tri les faces connectées
                    ListeFaceUsinageSection = TrierFacesConnectees(ListeFaceSection);

                    // Recherche des usinages d'extrémité et calcul des caractéristiques des usinages
                    // On calcul les distances des faces au points extremes
                    foreach (var l in ListeFaceUsinageSection)
                    {
                        l.CalculerUsinage(FaceSectionExt);
                        l.CalculerDistance(ExtremPoint1, ExtremPoint2);
                    }

                    // Recherche de la face la plus proche du point extreme 1
                    var Extrem = ListeFaceUsinageSection[0];
                    foreach (var l in ListeFaceUsinageSection)
                    {
                        if (Extrem.DistToExtremPoint1 > l.DistToExtremPoint1)
                            Extrem = l;
                    }

                    // On l'ajoute à la liste
                    ListeFaceUsinageExtremite.Add(Extrem);

                    // On recherche la face la plus proche du point extrème 2
                    // Elle peut être la même que la précédente
                    Extrem = ListeFaceUsinageSection[0];
                    foreach (var l in ListeFaceUsinageSection)
                    {
                        if (Extrem.DistToExtremPoint2 > l.DistToExtremPoint2)
                            Extrem = l;
                    }

                    // On l'ajoute
                    ListeFaceUsinageExtremite.AddIfNotExist(Extrem);

                    // On les supprime de la liste des faces de la section
                    foreach (var l in ListeFaceUsinageExtremite)
                        ListeFaceUsinageSection.Remove(l);
                }

                private List<ListFaceUsinage> TrierFacesConnectees(List<Face2> listeFace)
                {
                    List<ListFaceUsinage> ListeFacesUsinage = new List<ListFaceUsinage>();

                    if (listeFace.Count > 0)
                    {
                        List<Face2> ListeFaceTmp = new List<Face2>(listeFace);

                        // S'il y a des faces d'extremite non usinées
                        if (ListeFaceTmp.Count > 0)
                        {
                            ListeFacesUsinage.Add(new ListFaceUsinage(ListeFaceTmp[0]));
                            ListeFaceTmp.RemoveAt(0);

                            while (ListeFaceTmp.Count > 0)
                            {
                                var lst = ListeFacesUsinage.Last();

                                int i = 0;
                                while (i < ListeFaceTmp.Count)
                                {
                                    var f = ListeFaceTmp[i];

                                    if (lst.AjouterFaceConnectee(f))
                                    {
                                        ListeFaceTmp.RemoveAt(i);
                                        i = -1;
                                    }
                                    i++;
                                }

                                if (ListeFaceTmp.Count > 0)
                                {
                                    ListeFacesUsinage.Add(new ListFaceUsinage(ListeFaceTmp[0]));
                                    ListeFaceTmp.RemoveAt(0);
                                }
                            }
                        }
                    }

                    return ListeFacesUsinage;
                }

                public class ListFaceUsinage
                {
                    public Boolean Fermer = false;

                    public List<Face2> ListeFaces = new List<Face2>();
                    public List<Edge> ListeArretes = new List<Edge>();

                    /// <summary>
                    /// Liste des faces exterieur découpées
                    /// </summary>
                    public List<FaceGeom> ListeFaceDecoupe = new List<FaceGeom>();

                    /// <summary>
                    /// Liste des arretes des faces exterieures
                    /// </summary>
                    public List<Edge> ListeArreteDecoupe = new List<Edge>();

                    public Double LgUsinage = 0;

                    public Double DistToExtremPoint1 = 1E30;
                    public Double DistToExtremPoint2 = 1E30;

                    public void CalculerDistance(Point extremPoint1, Point extremPoint2)
                    {
                        foreach (var f in ListeFaces)
                        {
                            {
                                Double[] res = f.GetClosestPointOn(extremPoint1.X, extremPoint1.Y, extremPoint1.Z);
                                var dist = extremPoint1.Distance(new Point(res));
                                if (dist < DistToExtremPoint1) DistToExtremPoint1 = dist;
                            }

                            {
                                Double[] res = f.GetClosestPointOn(extremPoint2.X, extremPoint2.Y, extremPoint2.Z);
                                var dist = extremPoint2.Distance(new Point(res));
                                if (dist < DistToExtremPoint2) DistToExtremPoint2 = dist;
                            }
                        }
                    }

                    public void CalculerUsinage(ListFaceGeom faceExt)
                    {
                        Dictionary<FaceGeom, List<Edge>> ListeArreteExt = new Dictionary<FaceGeom, List<Edge>>();

                        foreach (var fg in faceExt.ListeFaceGeom)
                        {
                            var le = new List<Edge>();
                            var lf = new List<Face2>(fg.ListeSwFace);
                            ListeArreteExt.Add(fg, le);

                            foreach (var f in lf)
                                le.AddRange(f.eListeDesArretes());
                        }

                        foreach (var a in ListeArretes)
                        {
                            foreach (var fg in ListeArreteExt.Keys)
                            {
                                foreach (var ab in ListeArreteExt[fg])
                                {
                                    if (ab.eIsSame(a))
                                    {
                                        if(!ListeFaceDecoupe.Contains(fg))
                                            ListeFaceDecoupe.Add(fg);

                                        ListeArreteDecoupe.Add(a);
                                        LgUsinage += a.eLgArrete();
                                    }
                                }
                            }
                        }
                    }


                    // Initialisation avec une face
                    public ListFaceUsinage(Face2 f)
                    {
                        ListeFaces.Add(f);
                        ListeArretes.AddRange(f.eListeDesArretes());

                        // Verifie si la face est un cylindre
                        var cpt = 0;

                        foreach (var loop in f.eListeDesBoucles())
                            if (loop.IsOuter()) cpt++;

                        if (cpt > 1)
                            Fermer = true;
                    }

                    public Boolean AjouterFaceConnectee(Face2 f)
                    {
                        var result = UnionArretes(f.eListeDesArretes());

                        if (result > 0)
                        {
                            ListeFaces.Add(f);
                        }

                        if (result == 2)
                            Fermer = true;

                        return result > 0;
                    }

                    private Double UnionArretes(List<Edge> listeArretes)
                    {
                        var ListeTmp = new List<Edge>(listeArretes);
                        Double Connection = 0;

                        int i = 0;
                        while (i < ListeArretes.Count)
                        {
                            var Arrete1 = ListeArretes[i];

                            int j = 0;
                            while (j < ListeTmp.Count)
                            {
                                var Arrete2 = ListeTmp[j];

                                if (Arrete1.eIsSame(Arrete2))
                                {
                                    Connection++;

                                    ListeArretes.RemoveAt(i);
                                    ListeTmp.RemoveAt(j);
                                    i--;
                                    break;
                                }
                                j++;
                            }
                            i++;
                        }

                        if (Connection > 0)
                            ListeArretes.AddRange(ListeTmp);

                        return Connection;
                    }
                }

                #endregion

                #region ANALYSE DE LA GEOMETRIE ET RECHERCHE DU PROFIL

                private void AnalyserFaces()
                {
                    var ListeFaces = Corps.eListeDesFaces();

                    // On supprime les faces issues des fonctions appliquées
                    // sur la barre
                    foreach (var Fonc in Corps.eListeFonctions(null))
                    {
                        if (Fonc.GetTypeName2() != "WeldMemberFeat")
                        {
                            foreach (var f in Fonc.eListeDesFaces())
                                ListeFaces.Remove(f);
                        }
                    }

                    List<FaceGeom> ListeFaceExt = new List<FaceGeom>();

                    // Tri des faces pour retrouver celles issues de la même
                    foreach (var Face in ListeFaces)
                    {
                        var faceExt = new FaceGeom(Face);

                        Boolean Ajouter = true;

                        foreach (var f in ListeFaceExt)
                        {
                            // Si elles sont identiques, la face "faceExt" est ajoutée à la liste
                            // de face de "f"
                            if (f.FaceExtIdentique(faceExt))
                            {
                                Ajouter = false;
                                break;
                            }
                        }

                        // S'il n'y avait pas de face identique, on l'ajoute.
                        if (Ajouter)
                            ListeFaceExt.Add(faceExt);

                    }

                    // Analyse des faces et recherches du plan de section

                    var DicPlan = CombinerFaces(ListeFaceExt);

                    // On recherche le plan qui contient le plus de face
                    Plan Pmax = new Plan();
                    var ListeMax = new List<FaceGeom>();

                    foreach (var p in DicPlan.Keys)
                    {
                        var l = DicPlan[p];
                        var feBase = l[0];
                        if (l.Count > ListeMax.Count)
                        {
                            Pmax = p;
                            ListeMax = l;
                        }
                    }

                    // Plan de la section et infos
                    {
                        PlanSection = Pmax;
                        var v = PlanSection.Normale;
                        Double X = 0, Y = 0, Z = 0;
                        Corps.GetExtremePoint(v.X, v.Y, v.Z, out X, out Y, out Z);
                        ExtremPoint1 = new Point(X, Y, Z);
                        v.Inverser();
                        Corps.GetExtremePoint(v.X, v.Y, v.Z, out X, out Y, out Z);
                        ExtremPoint2 = new Point(X, Y, Z);
                    }

                    // Tri des faces section
                    ListeFaceSectionInt = TrierFacesConnectees(ListeMax);



                    // Tri des faces extremites
                    // On supprime les faces de la section
                    foreach (var f in ListeMax)
                        ListeFaceExt.Remove(f);

                    ListeFaceExtremite = TrierFacesConnectees(ListeFaceExt);

                    // On recherche la face exterieure
                    // s'il y a plusieurs boucles de surfaces
                    if (ListeFaceSectionInt.Count > 0)
                    {
                        {
                            // Si la section n'est composé que de cylindre fermé
                            Boolean EstUnCylindre = true;
                            ListFaceGeom Ext = null;
                            Double RayonMax = 0;
                            foreach (var fg in ListeFaceSectionInt)
                            {
                                if (fg.ListeFaceGeom.Count == 1)
                                {
                                    var f = fg.ListeFaceGeom[0];

                                    if (f.Type == eTypeFace.Cylindre)
                                    {
                                        if (RayonMax < f.Rayon)
                                        {
                                            RayonMax = f.Rayon;
                                            Ext = fg;
                                        }
                                    }
                                    else
                                    {
                                        EstUnCylindre = false;
                                        break;
                                    }
                                }
                            }

                            if (EstUnCylindre)
                            {
                                FaceSectionExt = Ext;
                                ListeFaceSectionInt.Remove(Ext);
                            }
                            else
                                FaceSectionExt = null;
                        }

                        {
                            // Methode plus longue pour determiner la face exterieur
                            if (FaceSectionExt == null)
                            {
                                // On créer un vecteur perpendiculaire à l'axe du profil
                                var vect = PlanSection.Normale;
                                vect = vect.Vectoriel(new Vecteur(1, 0, 0));
                                if (vect.X == 0)
                                    vect = vect.Vectoriel(new Vecteur(0, 0, 1));
                                vect.Normaliser();

                                // On récupère le point extreme dans cette direction
                                Double X = 0, Y = 0, Z = 0;
                                Corps.GetExtremePoint(vect.X, vect.Y, vect.Z, out X, out Y, out Z);
                                var Pt = new Point(X, Y, Z);

                                // La liste de face la plus proche est considérée comme la peau exterieur du profil
                                Double distMin = 1E30;
                                foreach (var Ext in ListeFaceSectionInt)
                                {
                                    foreach (var fg in Ext.ListeFaceGeom)
                                    {
                                        foreach (var f in fg.ListeSwFace)
                                        {
                                            Double[] res = f.GetClosestPointOn(Pt.X, Pt.Y, Pt.Z);
                                            var PtOnSurface = new Point(res);

                                            var dist = Pt.Distance(PtOnSurface);
                                            if (dist < 1E-6)
                                            {
                                                distMin = dist;
                                                FaceSectionExt = Ext;
                                                break;
                                            }
                                        }
                                    }
                                    if (FaceSectionExt.IsRef()) break;
                                }

                                // On supprime la face exterieur de la liste des faces
                                ListeFaceSectionInt.Remove(FaceSectionExt);
                            }
                        }
                    }
                    else
                    {
                        FaceSectionExt = ListeFaceSectionInt[0];
                    }
                }

                // FONCTION CRITIQUE
                private Dictionary<Plan, List<FaceGeom>> CombinerFaces(List<FaceGeom> listeFaceGeom)
                {
                    List<FaceGeom> ListeTest = new List<FaceGeom>(listeFaceGeom);

                    var DicPlan = new Dictionary<Plan, List<FaceGeom>>();

                    // On recherche les cylindre ou extrusion
                    foreach (var f in ListeTest)
                    {
                        if (f.Type == eTypeFace.Cylindre || f.Type == eTypeFace.Extrusion)
                        {
                            var plan = new Plan(f.Origine, f.Direction);

                            var Ajouter = true;
                            foreach (var p in DicPlan.Keys)
                            {
                                if (p.SontIdentiques(plan, 1E-10, false))
                                {
                                    DicPlan[p].Add(f);
                                    Ajouter = false;
                                }
                            }

                            if (Ajouter)
                                DicPlan.Add(plan, new List<FaceGeom>() { f });
                        }
                    }

                    FaceGeom depart = null;

                    // =================================================================================
                    // SECTION CRITIQUE
                    // =================================================================================
                    // Le choix de la face de depart conditionne le bon fonctionnement de la macro
                    // Il faut que cette face appartienne au profil et non à une extrémité
                    // =================================================================================
                    // 
                    // Choix de la face de départ
                    // S'il y a des faces cylindriques ou extrudées
                    // on les privilégie
                    // Sinon on part sur une face du milieu de la liste.
                    // Attention, erreur possible si la face de depart est une extrémité.

                    if (DicPlan.Count > 0)
                    {
                        var ltmp = new List<FaceGeom>();
                        foreach (var l in DicPlan.Values)
                        {
                            if (ltmp.Count < l.Count)
                                ltmp = l;
                        }

                        depart = ltmp[ltmp.Count / 2];
                    }
                    else
                    {
                        depart = ListeTest[ListeTest.Count / 2];
                    }
                    // ==================================================================================

                    // On récupère la liste des plans
                    List<FaceGeom> ListePlan = new List<FaceGeom>();
                    foreach (var f in ListeTest)
                    {
                        if (f.Type == eTypeFace.Plan)
                            ListePlan.Add(f);
                    }

                    if (ListePlan.Count > 2)
                    {
                        // On recherche les plans
                        foreach (var f in ListePlan)
                        {
                            var test = Orientation(f, depart);

                            if (test == eOrientation.Coplanaire || test == eOrientation.MemeOrigine)
                            {
                                // Dans le cas de vecteur normal parallèle, il faut ruser
                                var vF = (new Vecteur(depart.Origine, f.Origine)).Compose(f.Normale);
                                var v = depart.Normale.Vectoriel(vF);
                                var plan = new Plan(depart.Origine, v);

                                var Ajouter = true;
                                foreach (var p in DicPlan.Keys)
                                {
                                    if (p.SontIdentiques(plan, 1E-10, false))
                                    {
                                        DicPlan[p].Add(f);
                                        Ajouter = false;
                                    }
                                }

                                if (Ajouter)
                                    DicPlan.Add(plan, new List<FaceGeom>() { depart, f });
                            }
                        }

                        // On regarde si des points se retrouve sur des plans
                        foreach (var f in ListePlan)
                        {
                            var test = Orientation(f, depart);
                            if (test == eOrientation.Colineaire)
                            {
                                foreach (var p in DicPlan.Keys)
                                {
                                    if (p.SurLePlan(f.Origine, 1E-10))
                                    {
                                        DicPlan[p].Add(f);
                                    }
                                }
                            }
                        }


                    }

                    return DicPlan;
                }

                private List<ListFaceGeom> TrierFacesConnectees(List<FaceGeom> listeFace)
                {
                    List<FaceGeom> listeTmp = new List<FaceGeom>(listeFace);
                    List<ListFaceGeom> ListeTri = null;

                    if (listeTmp.Count > 0)
                    {
                        ListeTri = new List<ListFaceGeom>() { new ListFaceGeom(listeTmp[0]) };
                        listeTmp.RemoveAt(0);

                        while (listeTmp.Count > 0)
                        {
                            var l = ListeTri.Last();

                            int i = 0;
                            while (i < listeTmp.Count)
                            {
                                var f = listeTmp[i];

                                if (l.AjouterFaceConnectee(f))
                                {
                                    listeTmp.RemoveAt(i);
                                    i = -1;
                                }
                                i++;
                            }

                            if (listeTmp.Count > 0)
                            {
                                ListeTri.Add(new ListFaceGeom(listeTmp[0]));
                                listeTmp.RemoveAt(0);
                            }
                        }
                    }

                    // On recherche les cylindres uniques
                    // et on les marque comme fermé s'ils ont plus de deux boucle
                    foreach (var l in ListeTri)
                    {
                        if (l.ListeFaceGeom.Count == 1)
                        {
                            var f = l.ListeFaceGeom[0];
                            if (f.ListeSwFace.Count == 1)
                            {
                                var cpt = 0;

                                foreach (var loop in f.SwFace.eListeDesBoucles())
                                    if (loop.IsOuter()) cpt++;

                                if (cpt > 1)
                                    l.Fermer = true;
                            }
                        }
                    }

                    return ListeTri;
                }

                public enum eTypeFace
                {
                    Inconnu = 1,
                    Plan = 2,
                    Cylindre = 3,
                    Extrusion = 4
                }

                public enum eOrientation
                {
                    Indefini = 1,
                    Coplanaire = 2,
                    Colineaire = 3,
                    MemeOrigine = 4
                }

                public class FaceGeom
                {
                    public Face2 SwFace = null;
                    private Surface Surface = null;

                    public Point Origine;
                    public Vecteur Normale;
                    public Vecteur Direction;
                    public Double Rayon = 0;
                    public eTypeFace Type = eTypeFace.Inconnu;

                    public List<Face2> ListeSwFace = new List<Face2>();

                    public List<Face2> ListeFacesConnectee
                    {
                        get
                        {
                            var liste = new List<Face2>();

                            liste.AddRange(ListeSwFace[0].eListeDesFacesContigues());
                            for (int i = 1; i < ListeSwFace.Count; i++)
                            {
                                var l = ListeSwFace[i].eListeDesFacesContigues();

                                foreach (var f in l)
                                {
                                    liste.AddIfNotExist(f);
                                }

                            }

                            return liste;
                        }
                    }

                    public FaceGeom(Face2 swface)
                    {
                        SwFace = swface;

                        Surface = (Surface)SwFace.GetSurface();

                        ListeSwFace.Add(SwFace);

                        switch ((swSurfaceTypes_e)Surface.Identity())
                        {
                            case swSurfaceTypes_e.PLANE_TYPE:
                                Type = eTypeFace.Plan;
                                GetInfoPlan();
                                break;

                            case swSurfaceTypes_e.CYLINDER_TYPE:
                                Type = eTypeFace.Cylindre;
                                GetInfoCylindre();
                                break;

                            case swSurfaceTypes_e.EXTRU_TYPE:
                                Type = eTypeFace.Extrusion;
                                GetInfoExtrusion();
                                break;

                            default:
                                break;
                        }
                    }

                    public Boolean FaceExtIdentique(FaceGeom fe, Double arrondi = 1E-10)
                    {
                        if (Type != fe.Type)
                            return false;

                        if (!Origine.Comparer(fe.Origine, arrondi))
                            return false;

                        switch (Type)
                        {
                            case eTypeFace.Inconnu:
                                return false;
                            case eTypeFace.Plan:
                                if (!Normale.EstColineaire(fe.Normale, arrondi))
                                    return false;
                                break;
                            case eTypeFace.Cylindre:
                                if (!Direction.EstColineaire(fe.Direction, arrondi) || (Math.Abs(Rayon - fe.Rayon) > arrondi))
                                    return false;
                                break;
                            case eTypeFace.Extrusion:
                                if (!Direction.EstColineaire(fe.Direction, arrondi))
                                    return false;
                                break;
                            default:
                                break;
                        }

                        ListeSwFace.Add(fe.SwFace);
                        return true;
                    }

                    private void GetInfoPlan()
                    {
                        Boolean Reverse = SwFace.FaceInSurfaceSense();

                        if (Surface.IsPlane())
                        {
                            Double[] Param = Surface.PlaneParams;

                            if (Reverse)
                            {
                                Param[0] = Param[0] * -1;
                                Param[1] = Param[1] * -1;
                                Param[2] = Param[2] * -1;
                            }

                            Origine = new Point(Param[3], Param[4], Param[5]);
                            Normale = new Vecteur(Param[0], Param[1], Param[2]);
                        }
                    }

                    private void GetInfoCylindre()
                    {
                        if (Surface.IsCylinder())
                        {
                            Double[] Param = Surface.CylinderParams;

                            Origine = new Point(Param[0], Param[1], Param[2]);
                            Direction = new Vecteur(Param[3], Param[4], Param[5]);
                            Rayon = Param[6];

                            var UV = (Double[])SwFace.GetUVBounds();
                            Boolean Reverse = SwFace.FaceInSurfaceSense();

                            var ev1 = (Double[])Surface.Evaluate((UV[0] + UV[1]) * 0.5, (UV[2] + UV[3]) * 0.5, 0, 0);
                            if (Reverse)
                            {
                                ev1[3] = -ev1[3];
                                ev1[4] = -ev1[4];
                                ev1[5] = -ev1[5];
                            }

                            Normale = new Vecteur(ev1[3], ev1[4], ev1[5]);
                        }
                    }

                    private void GetInfoExtrusion()
                    {
                        if (Surface.IsSwept())
                        {
                            Double[] Param = Surface.GetExtrusionsurfParams();
                            Direction = new Vecteur(Param[0], Param[1], Param[2]);

                            Curve C = Surface.GetProfileCurve();
                            C = C.GetBaseCurve();

                            Double StartParam = 0, EndParam = 0;
                            Boolean IsClosed = false, IsPeriodic = false;

                            if (C.GetEndParams(out StartParam, out EndParam, out IsClosed, out IsPeriodic))
                            {
                                Double[] Eval = C.Evaluate(StartParam);

                                Origine = new Point(Eval[0], Eval[1], Eval[2]);
                            }

                            var UV = (Double[])SwFace.GetUVBounds();
                            Boolean Reverse = SwFace.FaceInSurfaceSense();

                            var ev1 = (Double[])Surface.Evaluate((UV[0] + UV[1]) * 0.5, (UV[2] + UV[3]) * 0.5, 0, 0);
                            if (Reverse)
                            {
                                ev1[3] = -ev1[3];
                                ev1[4] = -ev1[4];
                                ev1[5] = -ev1[5];
                            }

                            Normale = new Vecteur(ev1[3], ev1[4], ev1[5]);
                        }
                    }
                }

                public class ListFaceGeom
                {
                    public Boolean Fermer = false;

                    public List<FaceGeom> ListeFaceGeom = new List<FaceGeom>();

                    // Initialisation avec une face
                    public ListFaceGeom(FaceGeom f)
                    {
                        ListeFaceGeom.Add(f);
                    }

                    public List<Face2> ListeFaceSw()
                    {
                        var liste = new List<Face2>();

                        foreach (var fl in ListeFaceGeom)
                            liste.AddRange(fl.ListeSwFace);

                        return liste;
                    }

                    public Boolean AjouterFaceConnectee(FaceGeom f)
                    {
                        var Ajouter = false;
                        var Connection = 0;

                        int r = ListeFaceGeom.Count;

                        for (int i = 0; i < r; i++)
                        {
                            var l = ListeFaceGeom[i].ListeFacesConnectee;

                            foreach (var swf in f.ListeSwFace)
                            {
                                if (l.eContient(swf))
                                {
                                    if (Ajouter == false)
                                    {
                                        ListeFaceGeom.Add(f);
                                        Ajouter = true;
                                    }

                                    Connection++;
                                    break;
                                }
                            }

                        }

                        if (Connection > 1)
                            Fermer = true;

                        return Ajouter;
                    }
                }

                private eOrientation Orientation(FaceGeom f1, FaceGeom f2)
                {
                    var val = eOrientation.Indefini;
                    if (f1.Type == eTypeFace.Plan && f2.Type == eTypeFace.Plan)
                    {
                        val = Orientation(f1.Origine, f1.Normale, f2.Origine, f2.Normale);
                    }
                    else if (f1.Type == eTypeFace.Plan && (f2.Type == eTypeFace.Cylindre || f2.Type == eTypeFace.Extrusion))
                    {
                        Plan P = new Plan(f2.Origine, f2.Direction);
                        if (P.SurLePlan(f1.Origine, 1E-10) && P.SurLePlan(f1.Origine.Composer(f1.Normale), 1E-10))
                        {
                            val = eOrientation.Coplanaire;
                        }
                    }
                    else if (f2.Type == eTypeFace.Plan && (f1.Type == eTypeFace.Cylindre || f1.Type == eTypeFace.Extrusion))
                    {
                        Plan P = new Plan(f1.Origine, f1.Direction);
                        if (P.SurLePlan(f2.Origine, 1E-10) && P.SurLePlan(f2.Origine.Composer(f2.Normale), 1E-10))
                        {
                            val = eOrientation.Coplanaire;
                        }
                    }


                    return val;
                }

                private eOrientation Orientation(Point p1, Vecteur v1, Point p2, Vecteur v2)
                {
                    if (p1.Distance(p2) < 1E-10)
                        return eOrientation.MemeOrigine;

                    Vecteur Vtmp = new Vecteur(p1, p2);

                    if ((v1.Vectoriel(Vtmp).Norme < 1E-10) && (v2.Vectoriel(Vtmp).Norme < 1E-10))
                        return eOrientation.Colineaire;

                    Vecteur Vn1 = (new Vecteur(p1, p2)).Vectoriel(v1);
                    Vecteur Vn2 = (new Vecteur(p2, p1)).Vectoriel(v2);

                    Vecteur Vn = Vn1.Vectoriel(Vn2);

                    if (Vn.Norme < 1E-10)
                        return eOrientation.Coplanaire;

                    return eOrientation.Indefini;
                }

                #endregion
            }

            // ========================================================================================

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


