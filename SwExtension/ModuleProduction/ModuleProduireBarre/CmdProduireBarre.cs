using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.IO.Compression;

namespace ModuleProduction.ModuleProduireBarre
{
    public class CmdProduireBarre : Cmd
    {
        public ModelDoc2 MdlBase = null;
        public String RefFichier = "";
        public ListeSortedCorps ListeCorps = new ListeSortedCorps();
        public List<String> ListeMateriaux = new List<String>();
        public List<String> ListeProfils = new List<String>();
        public int Quantite = 1;
        public int IndiceCampagne = 0;
        public Boolean MettreAjourCampagne = false;
        public Boolean CreerPdf3D = false;
        public eTypeFichierExport TypeExport = eTypeFichierExport.ParasolidBinary;
        public Boolean ExporterBarres = true;
        public Boolean ListerUsinages = false;

        private InfosBarres Nomenclature = new InfosBarres();

        private String DossierBarre = "";
        private String DossierBarrePDF = "";
        private String NomDossierPDF = "PDF";
        private String NomFichierZip = "Laser-tube.zip";
        private String CheminFichierZip = "";

        private void Init()
        {
            CheminFichierZip = Path.Combine(DossierBarre, RefFichier + "_" + NomFichierZip);
            DossierBarre = Directory.CreateDirectory(Path.Combine(MdlBase.pDossierLaserTube(), IndiceCampagne.ToString())).FullName;
            if(CreerPdf3D) DossierBarrePDF = Directory.CreateDirectory(Path.Combine(DossierBarre, NomDossierPDF)).FullName;

            MdlBase.pCalculerQuantite(ref ListeCorps, eTypeCorps.Barre, ListeMateriaux, ListeProfils, IndiceCampagne, MettreAjourCampagne);
        }

        private void NettoyerFichier()
        {
            List<String> lstFichier = new List<String>();
            foreach (var f in Directory.GetFiles(DossierBarre))
            {
                if (Path.GetExtension(f) != ".txt")
                    lstFichier.Add(f);
            }

            if (MettreAjourCampagne)
            {
                List<String> ListeNomFichier = new List<string>();
                foreach (var corps in ListeCorps.Values)
                {
                    if (corps.Campagne[IndiceCampagne] > 0)
                    {
                        ListeNomFichier.Add(RefBarre(corps.RepereComplet, IndiceCampagne));
                    }
                }

                List<String> ListeFichierAsauvegarder = new List<string>();
                foreach (var nomFichier in ListeNomFichier)
                {
                    foreach (var f in lstFichier)
                    {
                        if (Path.GetFileNameWithoutExtension(f).StartsWith(nomFichier))
                        {
                            ListeFichierAsauvegarder.Add(f);
                            break;
                        }
                    }
                }

                foreach (var f in lstFichier)
                    if (!ListeFichierAsauvegarder.Contains(f))
                    {
                        File.Delete(f);
                        File.Delete(Path.Combine(DossierBarrePDF, Path.GetFileNameWithoutExtension(f) + ".pdf"));
                    }
            }
            else
            {
                File.Delete(Path.Combine(DossierBarre, CONST_PRODUCTION.FICHIER_NOMENC));

                foreach (var f in lstFichier)
                    File.Delete(f);

                if(Directory.Exists(DossierBarrePDF))
                    foreach (var f in Directory.GetFiles(DossierBarrePDF))
                        File.Delete(f);
            }

            File.Delete(CheminFichierZip);
        }

        protected override void Command()
        {
            try
            {
                Init();
                AffichageElementWPF Fenetre = new AffichageElementWPF(ListeCorps);
                Fenetre.OnValider += Start;
                Fenetre.ShowDialog();
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        private void Start()
        {
            try
            {
                NettoyerFichier();
                List<String> ListeFichiersZip = new List<string>();

                foreach (var corps in ListeCorps.Values)
                {
                    var chemin = CreerBarre(corps);
                    if(File.Exists(chemin))
                        ListeFichiersZip.Add(chemin);
                }

                WindowLog.SautDeLigne();
                WindowLog.Ecrire("Resumé :");
                foreach (var corps in ListeCorps.Values)
                    if (corps.Dvp && corps.Maj)
                        WindowLog.EcrireF("{2} P{0} ×{1}", corps.Repere, corps.Qte, IndiceCampagne);

                ListeCorps.EcrireProduction(DossierBarre, IndiceCampagne);

                if (ListerUsinages)
                    Nomenclature.TitreColonnes("Barre ref.", "Materiau", "Profil", "Lg", "Nb", "Usinage Ext 1", "Usinage Ext 2", "Détail des Usinage interne");
                else
                    Nomenclature.TitreColonnes("Barre ref.", "Materiau", "Profil", "Lg", "Nb");

                WindowLog.SautDeLigne();
                WindowLog.Ecrire(Nomenclature.ListeLignes());

                String FichierNomenclature_Laser = Path.Combine(DossierBarre, "Nomenclature_Laser.txt");
                StreamWriter s = new StreamWriter(FichierNomenclature_Laser);
                s.Write(Nomenclature.GenererTableau());
                s.Close();

                ListeFichiersZip.Add(FichierNomenclature_Laser);

                // Création du dossier ZIP
                if (ListeFichiersZip.Count > 1)
                {
                    WindowLog.SautDeLigne();
                    WindowLog.Ecrire(NomFichierZip);
                    WindowLog.Ecrire("  -> " + CheminFichierZip);

                    using (ZipArchive archive = ZipFile.Open(CheminFichierZip, ZipArchiveMode.Create))
                    {
                        foreach (var f in ListeFichiersZip)
                            archive.CreateEntryFromFile(f, Path.GetFileName(f), CompressionLevel.Optimal);

                        if(CreerPdf3D)
                            archive.CreateEntryFromFile(DossierBarrePDF, NomDossierPDF, CompressionLevel.Optimal);
                    }

                    foreach (var f in ListeFichiersZip)
                        File.Delete(f);

                    if (CreerPdf3D && Directory.Exists(DossierBarrePDF))
                    {
                        foreach (var f in Directory.GetFiles(DossierBarrePDF))
                            File.Delete(f);

                        Directory.Delete(DossierBarrePDF);
                    }
                }
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        private String CreerBarre(Corps corps)
        {
            string CheminFichierExportBarre = "";

            if (!corps.Dvp) return CheminFichierExportBarre;

            var QuantiteDiff = Quantite * (corps.Qte + corps.QteSup);

            var cheminFichier = corps.CheminFichierRepere;
            if (!File.Exists(cheminFichier)) return CheminFichierExportBarre;

            String Repere = corps.RepereComplet;
            String Profil = corps.Dimension;
            String Longueur = corps.Volume;
            String Materiau = corps.Materiau;
            String NomFichierBarre = ConstruireNomFichierBarre(Repere, IndiceCampagne, QuantiteDiff);

            var mdlCorps = Sw.eOuvrir(cheminFichier);
            if (mdlCorps.IsNull()) return CheminFichierExportBarre;

            var Piece = mdlCorps.ePartDoc();
            corps.SwCorps = Piece.ePremierCorps();
            Body2 Barre = corps.SwCorps;

            if(corps.Maj)
                WindowLog.EcrireF("{0}  x{1}", corps.RepereComplet, QuantiteDiff);

            mdlCorps.EditRebuild3();

            try
            {
                List<String> Liste = new List<String>() { Repere, Materiau, Profil, Math.Round(Longueur.eToDouble()).ToString(), "× " + QuantiteDiff.ToString() };

                if (ListerUsinages)
                {
                    var analyse = new AnalyseBarre(Barre, mdlCorps);

                    Dictionary<String, Double> Dic = new Dictionary<string, double>();

                    foreach (var u in analyse.ListeFaceUsinageSection)
                    {
                        String nom = u.ListeFaceDecoupe.Count + " face - Lg " + Math.Round(u.LgUsinage * 1000, 1);
                        if (Dic.ContainsKey(nom))
                            Dic[nom] += 1;
                        else
                            Dic.Add(nom, 1);
                    }

                    Liste.Add(Math.Round(analyse.ListeFaceUsinageExtremite[0].LgUsinage * 1000, 1).ToString());

                    if (analyse.ListeFaceUsinageExtremite.Count > 1)
                        Liste.Add(Math.Round(analyse.ListeFaceUsinageExtremite[1].LgUsinage * 1000, 1).ToString());
                    else
                        Liste.Add("");

                    foreach (var nom in Dic.Keys)
                        Liste.Add(Dic[nom] + "x [ " + nom + " ]");
                }

                Nomenclature.AjouterLigne(Liste.ToArray());

                if (corps.Maj && ExporterBarres)
                {
                    ModelDoc2 mdlBarre = Barre.eEnregistrerSous(Piece, DossierBarre, NomFichierBarre, TypeExport, out CheminFichierExportBarre);

                    if (CreerPdf3D)
                    {
                        String CheminPDF = Path.Combine(DossierBarrePDF, NomFichierBarre + eTypeFichierExport.PDF.GetEnumInfo<ExtFichier>());
                        mdlBarre.SauverEnPdf3D(CheminPDF);
                    }

                    App.Sw.CloseDoc(mdlBarre.GetPathName());
                }
            }
            catch (Exception e)
            {
                WindowLog.Ecrire("  - Erreur");
                this.LogMethode(new Object[] { e });
            }
            finally
            {
                WindowLog.Ecrire("  - Ok");
            }

            App.Sw.CloseDoc(mdlCorps.GetPathName());

            return CheminFichierExportBarre;
        }

        private class AnalyseBarre
        {
            public Body2 Corps = null;

            public ModelDoc2 Mdl = null;

            public gPlan PlanSection;
            public gPoint ExtremPoint1;
            public gPoint ExtremPoint2;
            public ListFaceGeom FaceSectionExt = null;
            public List<ListFaceGeom> ListeFaceSectionInt = null;
            public List<ListFaceUsinage> ListeFaceUsinageExtremite = new List<ListFaceUsinage>();
            public List<ListFaceUsinage> ListeFaceUsinageSection = new List<ListFaceUsinage>();

            public AnalyseBarre(Body2 corps, ModelDoc2 mdl)
            {
                Corps = corps;
                Mdl = mdl;

                AnalyserFaces();

                AnalyserPercages();
            }

            #region ANALYSE DES USINAGES

            private void AnalyserPercages()
            {
                // On recupère les faces issues des fonctions modifiant le corps
                List<Face2> ListeFaceSection = Corps.eListeDesFaces();

                // Supprimer les faces de la section ext
                foreach (var f in FaceSectionExt.ListeFaceSw())
                    ListeFaceSection.Remove(f);

                // Supprimer les faces des sections int
                foreach (var l in ListeFaceSectionInt)
                {
                    foreach (var f in l.ListeFaceSw())
                        ListeFaceSection.Remove(f);
                }

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

                public void CalculerDistance(gPoint extremPoint1, gPoint extremPoint2)
                {
                    foreach (var f in ListeFaces)
                    {
                        {
                            Double[] res = f.GetClosestPointOn(extremPoint1.X, extremPoint1.Y, extremPoint1.Z);
                            var dist = extremPoint1.Distance(new gPoint(res));
                            if (dist < DistToExtremPoint1) DistToExtremPoint1 = dist;
                        }

                        {
                            Double[] res = f.GetClosestPointOn(extremPoint2.X, extremPoint2.Y, extremPoint2.Z);
                            var dist = extremPoint2.Distance(new gPoint(res));
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
                                    ListeFaceDecoupe.AddIfNotExist(fg);
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
                try
                {
                    List<FaceGeom> ListeFaceCorps = new List<FaceGeom>();

                    // Tri des faces pour retrouver celles issues de la même
                    foreach (var Face in Corps.eListeDesFaces())
                    {
                        var faceExt = new FaceGeom(Face);

                        Boolean Ajouter = true;

                        foreach (var f in ListeFaceCorps)
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
                            ListeFaceCorps.Add(faceExt);

                    }

                    List<FaceGeom> ListeFaceGeom = new List<FaceGeom>();
                    PlanSection = RechercherFaceProfil(ListeFaceCorps, ref ListeFaceGeom);
                    ListeFaceSectionInt = TrierFacesConnectees(ListeFaceGeom);

                    // Plan de la section et infos
                    {
                        var v = PlanSection.Normale;
                        Double X = 0, Y = 0, Z = 0;
                        Corps.GetExtremePoint(v.X, v.Y, v.Z, out X, out Y, out Z);
                        ExtremPoint1 = new gPoint(X, Y, Z);
                        v.Inverser();
                        Corps.GetExtremePoint(v.X, v.Y, v.Z, out X, out Y, out Z);
                        ExtremPoint2 = new gPoint(X, Y, Z);
                    }

                    // =================================================================================

                    // On recherche la face exterieure
                    // s'il y a plusieurs boucles de surfaces
                    if (ListeFaceSectionInt.Count > 1)
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
                                var vect = this.PlanSection.Normale;

                                if (vect.X == 0)
                                    vect = vect.Vectoriel(new gVecteur(1, 0, 0));
                                else
                                    vect = vect.Vectoriel(new gVecteur(0, 0, 1));

                                vect.Normaliser();

                                // On récupère le point extreme dans cette direction
                                Double X = 0, Y = 0, Z = 0;
                                Corps.GetExtremePoint(vect.X, vect.Y, vect.Z, out X, out Y, out Z);
                                var Pt = new gPoint(X, Y, Z);

                                // La liste de face la plus proche est considérée comme la peau exterieur du profil
                                Double distMin = 1E30;
                                foreach (var Ext in ListeFaceSectionInt)
                                {
                                    foreach (var fg in Ext.ListeFaceGeom)
                                    {
                                        foreach (var f in fg.ListeSwFace)
                                        {
                                            Double[] res = f.GetClosestPointOn(Pt.X, Pt.Y, Pt.Z);
                                            var PtOnSurface = new gPoint(res);

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
                        ListeFaceSectionInt.RemoveAt(0);
                    }
                }
                catch (Exception e) { this.LogErreur(new Object[] { e }); }

            }

            private gPlan RechercherFaceProfil(List<FaceGeom> listeFaceGeom, ref List<FaceGeom> faceExt)
            {
                gPlan? p = null;
                try
                {
                    // On recherche les faces de la section
                    foreach (var fg in listeFaceGeom)
                    {
                        if (EstUneFaceProfil(fg))
                        {
                            faceExt.Add(fg);

                            // Si c'est un cylindre ou une extrusion, on recupère le plan
                            if ((p == null) && (fg.Type == eTypeFace.Cylindre || fg.Type == eTypeFace.Extrusion))
                                p = new gPlan(fg.Origine, fg.Direction);
                        }
                    }

                    // S'il n'y a que des faces plane, il faut calculer le plan de la section
                    // a partir de deux plan non parallèle
                    if (p == null)
                    {
                        gVecteur? v1 = null;
                        foreach (var fg in faceExt)
                        {
                            if (v1 == null)
                                v1 = fg.Normale;
                            else
                            {
                                var vtmp = ((gVecteur)v1).Vectoriel(fg.Normale);
                                if (Math.Abs(vtmp.Norme) > 1E-8)
                                    p = new gPlan(fg.Origine, vtmp);
                            }

                        }
                    }
                }
                catch (Exception e) { this.LogErreur(new Object[] { e }); }

                return (gPlan)p;
            }

            private Boolean EstUneFaceProfil(FaceGeom fg)
            {
                foreach (var f in fg.ListeSwFace)
                {
                    Byte[] Tab = Mdl.Extension.GetPersistReference3(f);
                    String S = System.Text.Encoding.Default.GetString(Tab);

                    int Pos_moSideFace = S.IndexOf("moSideFace3IntSurfIdRep_c");

                    int Pos_moVertexRef = S.Position("moVertexRef");

                    int Pos_moDerivedSurfIdRep = S.Position("moDerivedSurfIdRep_c");

                    int Pos_moFromSkt = Math.Min(S.Position("moFromSktEntSurfIdRep_c"), S.Position("moFromSktEnt3IntSurfIdRep_c"));

                    int Pos_moEndFace = Math.Min(S.Position("moEndFaceSurfIdRep_c"), S.Position("moEndFace3IntSurfIdRep_c"));

                    if (Pos_moSideFace != -1 && Pos_moSideFace < Pos_moEndFace && Pos_moSideFace < Pos_moFromSkt && Pos_moSideFace < Pos_moVertexRef && Pos_moSideFace < Pos_moDerivedSurfIdRep)
                        return true;
                }

                return false;
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

                public gPoint Origine;
                public gVecteur Normale;
                public gVecteur Direction;
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

                        Origine = new gPoint(Param[3], Param[4], Param[5]);
                        Normale = new gVecteur(Param[0], Param[1], Param[2]);
                    }
                }

                private void GetInfoCylindre()
                {
                    if (Surface.IsCylinder())
                    {
                        Double[] Param = Surface.CylinderParams;

                        Origine = new gPoint(Param[0], Param[1], Param[2]);
                        Direction = new gVecteur(Param[3], Param[4], Param[5]);
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

                        Normale = new gVecteur(ev1[3], ev1[4], ev1[5]);
                    }
                }

                private void GetInfoExtrusion()
                {
                    if (Surface.IsSwept())
                    {
                        Double[] Param = Surface.GetExtrusionsurfParams();
                        Direction = new gVecteur(Param[0], Param[1], Param[2]);

                        Curve C = Surface.GetProfileCurve();
                        C = C.GetBaseCurve();

                        Double StartParam = 0, EndParam = 0;
                        Boolean IsClosed = false, IsPeriodic = false;

                        if (C.GetEndParams(out StartParam, out EndParam, out IsClosed, out IsPeriodic))
                        {
                            Double[] Eval = C.Evaluate(StartParam);

                            Origine = new gPoint(Eval[0], Eval[1], Eval[2]);
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

                        Normale = new gVecteur(ev1[3], ev1[4], ev1[5]);
                    }
                }
            }

            public class ListFaceGeom
            {
                public Boolean Fermer = false;

                public List<FaceGeom> ListeFaceGeom = new List<FaceGeom>();

                public Double DistToExtremPoint1 = 1E30;
                public Double DistToExtremPoint2 = 1E30;

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

                public void CalculerDistance(gPoint extremPoint1, gPoint extremPoint2)
                {
                    foreach (var f in ListeFaceSw())
                    {
                        {
                            Double[] res = f.GetClosestPointOn(extremPoint1.X, extremPoint1.Y, extremPoint1.Z);
                            var dist = extremPoint1.Distance(new gPoint(res));
                            if (dist < DistToExtremPoint1) DistToExtremPoint1 = dist;
                        }

                        {
                            Double[] res = f.GetClosestPointOn(extremPoint2.X, extremPoint2.Y, extremPoint2.Z);
                            var dist = extremPoint2.Distance(new gPoint(res));
                            if (dist < DistToExtremPoint2) DistToExtremPoint2 = dist;
                        }
                    }
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
                    gPlan P = new gPlan(f2.Origine, f2.Direction);
                    if (P.SurLePlan(f1.Origine, 1E-10) && P.SurLePlan(f1.Origine.Composer(f1.Normale), 1E-10))
                    {
                        val = eOrientation.Coplanaire;
                    }
                }
                else if (f2.Type == eTypeFace.Plan && (f1.Type == eTypeFace.Cylindre || f1.Type == eTypeFace.Extrusion))
                {
                    gPlan P = new gPlan(f1.Origine, f1.Direction);
                    if (P.SurLePlan(f2.Origine, 1E-10) && P.SurLePlan(f2.Origine.Composer(f2.Normale), 1E-10))
                    {
                        val = eOrientation.Coplanaire;
                    }
                }


                return val;
            }

            private eOrientation Orientation(gPoint p1, gVecteur v1, gPoint p2, gVecteur v2)
            {
                if (p1.Distance(p2) < 1E-10)
                    return eOrientation.MemeOrigine;

                gVecteur Vtmp = new gVecteur(p1, p2);

                if ((v1.Vectoriel(Vtmp).Norme < 1E-10) && (v2.Vectoriel(Vtmp).Norme < 1E-10))
                    return eOrientation.Colineaire;

                gVecteur Vn1 = (new gVecteur(p1, p2)).Vectoriel(v1);
                gVecteur Vn2 = (new gVecteur(p2, p1)).Vectoriel(v2);

                gVecteur Vn = Vn1.Vectoriel(Vn2);

                if (Vn.Norme < 1E-10)
                    return eOrientation.Coplanaire;

                return eOrientation.Indefini;
            }

            #endregion
        }

        // ========================================================================================

        public String RefBarre(String refBarre, int campagne)
        {
            return String.Format("{0}{1}-{2}", String.IsNullOrWhiteSpace(RefFichier) ? "" : RefFichier + "-", campagne, refBarre);
        }

        public String ConstruireNomFichierBarre(String refBarre, int campagne, int quantite)
        {
            return String.Format("{0} (×{1})", RefBarre(refBarre, campagne), quantite);
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

            public void AjouterLigne(String[] Valeurs)
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

            public SortedSet<String> ListeLignes()
            {
                SortedSet<String> Liste = new SortedSet<String>(new WindowsStringComparer());

                try
                {
                    if (_TitreColonnes.Count > 0)
                    {
                        String formatTitre = "";

                        for (int i = 0; i < _TitreColonnes.Count; i++)
                            formatTitre += "{" + i.ToString() + ",-" + _DimColonnes[i] + "}    ";

                        formatTitre = formatTitre.Trim();

                        Liste.Add(String.Format(formatTitre, _TitreColonnes.ToArray()));
                    }

                    if (Count > 0)
                    {
                        String format = "";

                        for (int i = 0; i < _DimColonnes.Count; i++)
                        {
                            String Sign = "";
                            var l = this[0];

                            if (i < l.Count && l[i].Trim().Count() > 0 && Char.IsLetter(l[i].Trim()[0]))
                                Sign = "-";

                            format += "{" + i.ToString() + "," + Sign + _DimColonnes[i] + "}    ";
                        }

                        format = format.Trim();

                        foreach (List<String> ligne in this)
                        {
                            String[] t = new String[_DimColonnes.Count];

                            for (int i = 0; i < _DimColonnes.Count; i++)
                            {
                                if (i < ligne.Count)
                                    t[i] = ligne[i];
                                else
                                    t[i] = "";
                            }


                            Liste.Add(String.Format(format, t));

                        }
                    }
                }
                catch (Exception e)
                { this.LogErreur(new Object[] { e }); }

                return Liste;
            }

            public String GenererTableau()
            {
                return String.Join("\r\n", ListeLignes());
            }
        }
    }
}


