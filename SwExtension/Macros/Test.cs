using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Test"),
        ModuleNom("Test")]

    public class Test : BoutonBase
    {
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
                    foreach (var f in Fonc.eListeDesFaces())
                    {
                        Body2 cF = f.GetBody();
                        if (cF.eIsSame(Corps))
                            ListeFaceSection.AddIfNotExist(f);
                    }
                }

                foreach (var f in FaceSectionExt.ListeFaceSw())
                    ListeFaceSection.Remove(f);

                foreach (var l in ListeFaceSectionInt)
                {
                    foreach (var f in l.ListeFaceSw())
                        ListeFaceSection.Remove(f);
                }

                foreach (var l in ListeFaceExtremite)
                {
                    foreach (var f in l.ListeFaceSw())
                        ListeFaceSection.Remove(f);
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
                    var ListeFaces = Corps.eListeDesFaces();

                    // On supprime les faces issues des fonctions appliquées
                    // sur la barre
                    foreach (var Fonc in Corps.eListeFonctions(null))
                    {
                        if (!Fonc.IsBase2())
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

                    foreach (var l in ListeFaceSectionInt)
                    {
                        foreach (var f in l.ListeFaceSw())
                        {
                            f.eSelectEntite(true);
                        }
                    }

                    // Tri des faces extremites
                    // On supprime les faces de la section
                    foreach (var f in ListeMax)
                        ListeFaceExt.Remove(f);

                    // =================================================================================

                    ListeFaceExtremite = TrierFacesConnectees(ListeFaceExt);

                    var lstTmp = new List<ListFaceGeom>();
                    lstTmp.AddRange(ListeFaceExtremite);
                    ListeFaceExtremite.Clear();

                    // Recherche des faces d'extrémité
                    // On calcul les distances des faces au points extremes
                    foreach (var l in lstTmp)
                        l.CalculerDistance(ExtremPoint1, ExtremPoint2);

                    // Recherche de la face la plus proche du point extreme 1
                    var Extrem = lstTmp[0];
                    foreach (var l in lstTmp)
                    {
                        if (Extrem.DistToExtremPoint1 > l.DistToExtremPoint1)
                            Extrem = l;
                    }

                    // On l'ajoute à la liste
                    ListeFaceExtremite.Add(Extrem);

                    // On recherche la face la plus proche du point extrème 2
                    // Elle peut être la même que la précédente
                    Extrem = lstTmp[0];
                    foreach (var l in lstTmp)
                    {
                        if (Extrem.DistToExtremPoint2 > l.DistToExtremPoint2)
                            Extrem = l;
                    }

                    // On l'ajoute
                    ListeFaceExtremite.AddIfNotExist(Extrem);

                    // =================================================================================

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
                                if (vect.X == 0)
                                    vect = vect.Vectoriel(new Vecteur(1, 0, 0));
                                else
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
                catch (Exception e) { this.LogMethode(new Object[] { e }); }

            }

            // FONCTION CRITIQUE
            private Dictionary<Plan, List<FaceGeom>> CombinerFaces(List<FaceGeom> listeFaceGeom)
            {
                List<FaceGeom> ListeTest = new List<FaceGeom>(listeFaceGeom);

                var DicPlan = new Dictionary<Plan, List<FaceGeom>>();

                try
                {
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
                }
                catch (Exception e) { this.LogMethode(new Object[] { e }); }

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

                public void CalculerDistance(Point extremPoint1, Point extremPoint2)
                {
                    foreach (var f in ListeFaceSw())
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

        protected override void Command()
        {
            try
            {
                ModelDoc2 mdl = App.ModelDoc2;
                //var DossierExport = mdl.eDossier();
                //var NomFichier = mdl.eNomSansExt();

                Body2 Barre = null;

                var Face = mdl.eSelect_RecupererObjet<Face2>(1);

                if (Face.IsNull())
                    Barre = mdl.eSelect_RecupererObjet<Body2>(1);
                else
                    Barre = Face.GetBody();

                mdl.eEffacerSelection();

                WindowLog.Ecrire("Nom du corps : " + Barre.Name);

                var b = new AnalyseBarre(Barre);

                //foreach (var u in b.ListeFaceUsinageExtremite)
                //{
                //    WindowLog.Ecrire(u.LgUsinage * 1000);
                //    foreach (var e in u.ListeArreteDecoupe)
                //        e.eSelectEntite(true);
                //}

                foreach (var u in b.ListeFaceUsinageSection)
                {
                    WindowLog.Ecrire(u.LgUsinage * 1000);
                    foreach (var e in u.ListeArreteDecoupe)
                        e.eSelectEntite(true);
                }
            }
            catch (Exception e) { this.LogMethode(new Object[] { e }); }

        }
    }

    //public class Test : BoutonBase
    //{
    //    protected override void Command()
    //    {
    //        try
    //        {
    //            ModelDoc2 mdl = App.ModelDoc2;
    //            var DossierExport = mdl.eDossier();
    //            var NomFichier = mdl.eNomSansExt();

    //            var ListeNomConfigs = mdl.eListeNomConfiguration(eTypeConfig.Pliee);
    //            ListeNomConfigs.Sort(new WindowsStringComparer());

    //            for (int noCfg = 0; noCfg < ListeNomConfigs.Count; noCfg++)
    //            {
    //                mdl.ClearSelection2(true);

    //                var NomConfigPliee = ListeNomConfigs[noCfg];
    //                mdl.ShowConfiguration2(NomConfigPliee);
    //                mdl.EditRebuild3();
    //                PartDoc Piece = mdl.ePartDoc();

    //                ListPID<Feature> ListeDossier = Piece.eListePIDdesFonctionsDePiecesSoudees(null);

    //                for (int noD = 0; noD < ListeDossier.Count; noD++)
    //                {
    //                    Feature f = ListeDossier[noD];
    //                    BodyFolder dossier = f.GetSpecificFeature2();

    //                    if (dossier.eEstExclu() || dossier.IsNull() || (dossier.GetBodyCount() == 0)) continue;

    //                    String Profil = dossier.eProp(CONSTANTES.PROFIL_NOM);
    //                    String Longueur = dossier.eProp(CONSTANTES.PROFIL_LONGUEUR);

    //                    if (String.IsNullOrWhiteSpace(Profil) || String.IsNullOrWhiteSpace(Longueur))
    //                    {
    //                        WindowLog.Ecrire("      Pas de barres");
    //                        continue;
    //                    }

    //                    foreach (var Barre in dossier.eListeDesCorps())
    //                        Barre.Select2(true, null);
    //                }

    //                var mdlExport = ExportSelection(Piece, DossierExport, NomFichier + "-" + noCfg + "-Export Tube", eTypeFichierExport.Piece);

    //                mdlExport.ViewZoomtofit2();
    //                mdlExport.ShowNamedView2("*Isométrique", 7);
    //                int lErrors = 0, lWarnings = 0;
    //                mdlExport.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref lErrors, ref lWarnings);

    //                App.Sw.CloseDoc(mdlExport.GetPathName());
    //            }

    //        }
    //        catch (Exception e) { this.LogMethode(new Object[] { e }); }

    //    }

    //    private ModelDoc2 ExportSelection(PartDoc piece, String dossier, String nomFichier, eTypeFichierExport typeExport)
    //    {
    //        int pStatut;
    //        int pWarning;

    //        Boolean Resultat = piece.SaveToFile3(Path.Combine(dossier, nomFichier + typeExport.GetEnumInfo<ExtFichier>()),
    //                                              (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
    //                                              (int)swCutListTransferOptions_e.swCutListTransferOptions_CutListProperties,
    //                                              false,
    //                                              "",
    //                                              out pStatut,
    //                                              out pWarning);
    //        if (Resultat)
    //            return App.ModelDoc2;

    //        return null;
    //    }
    //}

    //public class Test : BoutonBase
    //{
    //    protected override void Command()
    //    {
    //        try
    //        {
    //            ModelDoc2 MdlBase = App.ModelDoc2;

    //            Feature Fonction = MdlBase.eSelect_RecupererObjet<Feature>(1, -1);

    //            MdlBase.eEffacerSelection();

    //            Sketch Esquisse = Fonction.GetSpecificFeature2();

    //            List<Point> ListPt = new List<Point>();

    //            HashSet<String> ListId = new HashSet<string>();

    //            Func<SketchPoint, String> IdString = delegate (SketchPoint sp)
    //            {
    //                int[] id = (int[])sp.GetID();
    //                return id[0] + "-" + id[1];
    //            };


    //            foreach (SketchSegment sg in Esquisse.GetSketchSegments())
    //            {
    //                if (sg.GetType() != (int)swSketchSegments_e.swSketchLINE)
    //                    continue;

    //                SketchLine l = sg as SketchLine;
    //                SketchPoint pt;

    //                pt = l.GetStartPoint2();
    //                if (!ListId.Contains(IdString(pt)))
    //                    ListPt.Add(new Point(pt));

    //                pt = l.GetEndPoint2();
    //                if (!ListId.Contains(IdString(pt)))
    //                    ListPt.Add(new Point(pt));
    //            }

    //            if (ListPt.Count > 0)
    //            {
    //                String Fichier = Path.Combine(MdlBase.eDossier(), "ExportPoint.csv");

    //                using (StreamWriter Sw = File.CreateText(Fichier))
    //                {
    //                    foreach (var pt in ListPt)
    //                    {
    //                        Sw.WriteLine(String.Format("{0};{1};{2}", Math.Round(pt.X, 3), Math.Round(pt.Y, 3), Math.Round(pt.Z, 3)));
    //                        WindowLog.EcrireF("{0} {1} {2}", Math.Round(pt.X, 3), Math.Round(pt.Y, 3), Math.Round(pt.Z, 3));
    //                    }
    //                }
    //            }
    //        }
    //        catch (Exception e) { this.LogMethode(new Object[] { e }); }

    //    }
    //}

    //public class TestOld : BoutonBase
    //{

    //    private List<String> CalquesBase = new List<string>() { "Annotations", "Cotations", "Tables", "Vue", "Construction", "Bordure", "Pliage" };

    //    protected override void Command()
    //    {
    //        try
    //        {
    //            int lErrors = 0;
    //            int lWarnings = 0;

    //            String cheminDossier = App.Sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsSheetFormat);

    //            foreach (var cheminFichier in Directory.GetFiles(cheminDossier))
    //            {

    //                ModelDoc2 MdlBase = App.Sw.OpenDoc6(cheminFichier, (int)swDocumentTypes_e.swDocDRAWING, 0, "", ref lErrors, ref lWarnings);

    //                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swWeldmentEnableAutomaticCutList, 0, true);
    //                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swWeldmentEnableAutomaticUpdate, 0, false);
    //                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisableDerivedConfigurations, 0, false);
    //                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swWeldmentRenameCutlistDescriptionPropertyValue, 0, true);
    //                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swWeldmentCollectIdenticalBodies, 0, true);
    //                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSheetMetalBodiesDescriptionUseDefault, 0, false);
    //                MdlBase.Extension.SetUserPreferenceString((int)swUserPreferenceStringValue_e.swSheetMetalDescription, 0, "Tôle");
    //                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swFlatPatternOpt_SimplifyBends, 0, false);
    //                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swFlatPatternOpt_CornerTreatment, 0, false);

    //                if (MdlBase.TypeDoc() == eTypeDoc.Dessin)
    //                {

    //                    LayerMgr LM = MdlBase.GetLayerManager();

    //                    String[] ListeCalques = LM.GetLayerList();

    //                    WindowLog.Ecrire(MdlBase.GetPathName());

    //                    foreach (var Calque in ListeCalques)
    //                    {
    //                        if (!CalquesBase.Contains(Calque))
    //                        {
    //                            WindowLog.Ecrire(Calque + " : " + LM.DeleteLayer(Calque));
    //                        }
    //                    }

    //                    String cheminFondPlan = MdlBase.eDrawingDoc().eFeuilleActive().eGetGabaritDeFeuille();

    //                    String nomFondPlan = cheminFondPlan.Replace(cheminDossier + "\\", "");

    //                    WindowLog.Ecrire(nomFondPlan);

    //                    //MdlBase.Extension.DeleteDraftingStandard();

    //                    //MdlBase.ForceRebuild3(false);

    //                    if (nomFondPlan.ToLower().StartsWith("archi"))
    //                    {
    //                        Boolean r = MdlBase.Extension.LoadDraftingStandard("E:\\Mes documents\\SolidWorks\\2018\\Norme dessin\\Norme Archi.sldstd");
    //                        WindowLog.Ecrire("Norme Archi.sldstd" + " : " + r);
    //                    }
    //                    else
    //                    {
    //                        Boolean r = MdlBase.Extension.LoadDraftingStandard("E:\\Mes documents\\SolidWorks\\2018\\Norme dessin\\Norme Fab.sldstd");
    //                        WindowLog.Ecrire("Norme Fab.sldstd" + " : " + r);
    //                    }

    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitSystem, 0, (int)swUnitSystem_e.swUnitSystem_Custom);
    //                    MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swUnitsLinearFeetAndInchesFormat, 0, false);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsDualLinearFractionDenominator, 0, 0);
    //                    MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swUnitsDualLinearFeetAndInchesFormat, 0, false);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swSheetMetalColorFlatPatternSketch, 0, 8421504);
    //                    TextFormat myTextFormat = MdlBase.Extension.GetUserPreferenceTextFormat((int)swUserPreferenceTextFormat_e.swSheetMetalBendNotesTextFormat, 0);
    //                    myTextFormat.CharHeight = 0.004;
    //                    MdlBase.Extension.SetUserPreferenceTextFormat((int)swUserPreferenceTextFormat_e.swSheetMetalBendNotesTextFormat, 0, myTextFormat);
    //                    MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSheetMetalBendNotesLeaderJustificationSnapping, 0, true);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDetailingAltLinearDimPrecision, (int)swUserPreferenceOption_e.swDetailingLinearDimension, 2);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDetailingAltLinearDimPrecision, (int)swUserPreferenceOption_e.swDetailingDiameterDimension, 2);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDimensionsExtensionLineStyle, (int)swUserPreferenceOption_e.swDetailingRadiusDimension, (int)swLineStyles_e.swLineCONTINUOUS);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDimensionsExtensionLineStyleThickness, (int)swUserPreferenceOption_e.swDetailingRadiusDimension, (int)swLineWeights_e.swLW_THIN);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDetailingAltLinearDimPrecision, (int)swUserPreferenceOption_e.swDetailingRadiusDimension, 2);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDimensionsExtensionLineStyle, (int)swUserPreferenceOption_e.swDetailingHoleDimension, (int)swLineStyles_e.swLineCONTINUOUS);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDimensionsExtensionLineStyleThickness, (int)swUserPreferenceOption_e.swDetailingHoleDimension, (int)swLineWeights_e.swLW_THIN);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDetailingLinearDimPrecision, (int)swUserPreferenceOption_e.swDetailingHoleDimension, 4);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDetailingAltLinearDimPrecision, (int)swUserPreferenceOption_e.swDetailingHoleDimension, 2);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDetailingAngleTrailingZero, (int)swUserPreferenceOption_e.swDetailingAngleDimension, (int)swDetailingDimTrailingZero_e.swDimRemoveTrailingZeroes);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDetailingAngleTrailingZeroTolerance, (int)swUserPreferenceOption_e.swDetailingAngleDimension, (int)swDetailingDimTrailingZero_e.swDimSameAsDocumentTolerance);
    //                    MdlBase.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDimensionsExtensionLineStyle, (int)swUserPreferenceOption_e.swDetailingChamferDimension, (int)swLineStyles_e.swLineCONTINUOUS);

    //                    //MdlBase.ForceRebuild3(false);

    //                    MdlBase.eDrawingDoc().eFeuilleActive().SaveFormat(cheminFondPlan);
    //                }
    //                else
    //                {
    //                    MdlBase.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref lErrors, ref lWarnings);
    //                }

    //                App.Sw.CloseDoc(MdlBase.GetPathName());
    //            }



    //        }
    //        catch (Exception e) { this.LogMethode(new Object[] { e }); }

    //    }
    //}
}
