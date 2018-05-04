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
        private class Barre
        {
            public Body2 Corps = null;

            public List<List<Face2>> Liste = new List<List<Face2>>();

            public Barre(Body2 corps)
            {
                Corps = corps;
            }

            private class FaceExt
            {
                public Face2 SwFace = null;

                public Vecteur Normale;
                public Vecteur Vectoriel;

                public Dictionary<Face2, Vecteur> ListeVect = new Dictionary<Face2, Vecteur>();
                public FaceExt(Face2 swface)
                {
                    SwFace = swface;

                    Normale = GetNormal(SwFace);

                    foreach (var e in SwFace.eListeDesArretes())
                    {
                        var c = (Curve)e.GetCurve();
                        if (c.IsLine())
                        {
                            var l = e.eListeDesFaces();
                            l.Remove(SwFace);
                            var coFace = l[0];
                            var vec = GetNormal(coFace);
                        }
                    }
                }

                private Vecteur GetNormal(Face2 face)
                {
                    Vecteur V = new Vecteur();

                    Boolean Reverse = face.FaceInSurfaceSense();
                    var S = (Surface)face.GetSurface();

                    switch ((swSurfaceTypes_e)S.Identity())
                    {
                        case swSurfaceTypes_e.PLANE_TYPE:
                            {
                                WindowLog.Ecrire("Plan");
                                Double[] Param = S.PlaneParams;

                                if (Reverse)
                                {
                                    Param[0] = Param[0] * -1;
                                    Param[1] = Param[1] * -1;
                                    Param[2] = Param[2] * -1;
                                }

                                V = new Vecteur(Param[0], Param[1], Param[2]);
                            }
                            break;
                        case swSurfaceTypes_e.CYLINDER_TYPE:
                            {
                                WindowLog.Ecrire("Cylindre");

                                var UV = (Double[])face.GetUVBounds();
                                var ev = (Double[])S.Evaluate((UV[0] + UV[1]) * 0.5, (UV[2] + UV[3]) * 0.5, 0, 0);

                                if (Reverse)
                                {
                                    ev[3] = -ev[3];
                                    ev[4] = -ev[4];
                                    ev[5] = -ev[5];
                                }

                                V = new Vecteur(ev[3], ev[4], ev[5]);
                            }
                            break;
                        default:
                            break;
                    }

                    return V;
                }

                public void Analyser()
                {

                }

                //public void Analyser()
                //{
                //    var ListeFaces = Corps.eListeDesFaces();
                //    var ListeFonc = Corps.eListeFonctions(null, false);
                //    foreach (var F in ListeFonc)
                //    {
                //        WindowLog.Ecrire(F.Name + " -> " + F.GetTypeName2());
                //        if (F.GetTypeName2() != "WeldMemberFeat")
                //        {
                //            foreach (var f in F.eListeDesFaces())
                //            {
                //                ListeFaces.Remove(f);
                //            }
                //        }

                //        WindowLog.Ecrire(F.Name + " -> " + F.GetTypeName2());
                //    }

                //    var SM = Sw.eModeleActif().SketchManager;

                //    FaceBarre LastFace = null;

                //    var ListeTri = new List<List<FaceBarre>>();
                //    List<FaceBarre> ListeCourante = new List<FaceBarre>();

                //    foreach (var Face in ListeFaces)
                //    {
                //        var B = (Body2)Face.GetBody();
                //        if (B.Name == Corps.Name)
                //        {
                //            Surface S = Face.GetSurface();

                //            Boolean Reverse = Face.FaceInSurfaceSense();
                //            Double[] Point = new Double[3];
                //            Double[] Normale = new Double[3];
                //            switch ((swSurfaceTypes_e)S.Identity())
                //            {
                //                case swSurfaceTypes_e.PLANE_TYPE:
                //                    {
                //                        WindowLog.Ecrire("Plan");
                //                        Double[] Param = S.PlaneParams;

                //                        if (Reverse)
                //                        {
                //                            Param[0] = Param[0] * -1;
                //                            Param[1] = Param[1] * -1;
                //                            Param[2] = Param[2] * -1;
                //                        }

                //                        Point[0] = Param[3]; Point[1] = Param[4]; Point[2] = Param[5];
                //                        Normale[0] = Param[0]; Normale[1] = Param[1]; Normale[2] = Param[2];
                //                    }
                //                    break;
                //                case swSurfaceTypes_e.CYLINDER_TYPE:
                //                    {
                //                        WindowLog.Ecrire("Cylindre");
                //                        Double[] Param = S.CylinderParams;

                //                        Point[0] = Param[0]; Point[1] = Param[1]; Point[2] = Param[2];
                //                        Normale = NormaleCylindre(new Double[] { Param[3], Param[4], Param[5] });
                //                    }
                //                    break;
                //                case swSurfaceTypes_e.CONE_TYPE:
                //                    break;
                //                case swSurfaceTypes_e.SPHERE_TYPE:
                //                    break;
                //                case swSurfaceTypes_e.TORUS_TYPE:
                //                    break;
                //                case swSurfaceTypes_e.BSURF_TYPE:
                //                    break;
                //                case swSurfaceTypes_e.BLEND_TYPE:
                //                    break;
                //                case swSurfaceTypes_e.OFFSET_TYPE:
                //                    break;
                //                case swSurfaceTypes_e.EXTRU_TYPE:
                //                    break;
                //                case swSurfaceTypes_e.SREV_TYPE:
                //                    break;
                //                default:
                //                    break;
                //            }

                //            //SM.Insert3DSketch(false);
                //            //SM.AddToDB = true;
                //            //SM.DisplayWhenAdded = false;

                //            //SM.CreatePoint(Point[0], Point[1], Point[2]);
                //            //SM.CreateLine(Point[0], Point[1], Point[2], Point[0] + (Normale[0] * 0.001), Point[1] + (Normale[1] * 0.001), Point[2] + (Normale[2] * 0.001));

                //            //SM.DisplayWhenAdded = true;
                //            //SM.AddToDB = false;
                //            //SM.Insert3DSketch(true);

                //            var faceBarre = new FaceBarre(Point, Normale, Face);

                //            if (LastFace.IsRef() && faceBarre.EstCoplanaire(LastFace))
                //            {
                //                ListeCourante.Add(faceBarre);
                //            }
                //            else
                //            {
                //                var Stop = false;
                //                foreach (var L in ListeTri)
                //                {
                //                    foreach (var fb in L)
                //                    {
                //                        if (fb.MemeOrigine(faceBarre))
                //                        {
                //                            L.Add(faceBarre);
                //                            ListeCourante = L;
                //                            Stop = true;
                //                            break;
                //                        }
                //                    }
                //                    if (Stop) break;
                //                }

                //                if (!Stop)
                //                {
                //                    ListeCourante = new List<FaceBarre>();
                //                    ListeCourante.Add(faceBarre);
                //                    ListeTri.Add(ListeCourante);
                //                }
                //            }

                //            LastFace = faceBarre;
                //        }

                //    }

                //    foreach (var liste in ListeTri)
                //    {
                //        var l = new List<Face2>();
                //        Liste.Add(l);
                //        foreach (var f in liste)
                //            l.Add(f.Face);
                //    }
                //}

                //private class FaceBarre
                //{
                //    public Double[] Point = new Double[] { 0, 0, 0 };
                //    public Double[] Normale = new Double[] { 0, 0, 0 };
                //    public Face2 Face = null;

                //    public FaceBarre(Double[] point, Double[] normale, Face2 face)
                //    {
                //        Point = point; Normale = normale; Face = face;
                //    }

                //    public Boolean MemeOrigine(FaceBarre faceBarre)
                //    {
                //        var Pt = faceBarre.Point;
                //        var Nl = faceBarre.Normale;

                //        if (Pt[0] == Point[0] && Pt[1] == Point[1] && Pt[2] == Point[2] && Nl[0] == Normale[0] && Nl[1] == Normale[1] && Nl[2] == Normale[2])
                //            return true;

                //        return false;
                //    }

                //    public Boolean EstConnecte(FaceBarre f)
                //    {
                //        return Face.eFaceEstConnecte(f.Face);
                //    }

                //    public Boolean EstCoplanaire(FaceBarre f)
                //    {
                //        return EstCoplanaire(Point, Normale, f.Point, f.Normale);
                //    }

                //    private Boolean EstCoplanaire(Double[] pt1, Double[] vec1, Double[] pt2, Double[] vec2)
                //    {

                //        Point P1 = new Point(pt1);
                //        Vecteur V1 = new Vecteur(vec1);
                //        Point P2 = new Point(pt2);
                //        Vecteur V2 = new Vecteur(vec2);

                //        if (P1.Distance(P2) < 1E-10)
                //            return true;

                //        Vecteur Vtmp = new Vecteur(P1, P2);

                //        if ((V1.Vectoriel(Vtmp).Norme < 1E-10) && (V2.Vectoriel(Vtmp).Norme < 1E-10))
                //            return true;

                //        Vecteur Vn1 = (new Vecteur(P1, P2)).Vectoriel(V1);
                //        Vecteur Vn2 = (new Vecteur(P2, P1)).Vectoriel(V2);

                //        Vecteur Vn = Vn1.Vectoriel(Vn2);

                //        if (Vn.Norme < 1E-10)
                //            return true;

                //        return false;
                //    }
                //}

                //private Double[] NormaleCylindre(Double[] Axe)
                //{
                //    Double[] Normale = new Double[] { 0, 0, 0 };

                //    if (Axe[0] == 0 && Axe[1] == 0)
                //        Normale[0] = 1;
                //    else
                //    {
                //        Normale[0] = Axe[1];
                //        Normale[1] = -1 * Axe[0];
                //    }

                //    return Normale;
                //}

            }
        }

        protected override void Command()
        {
            try
            {
                ModelDoc2 mdl = App.ModelDoc2;
                //var DossierExport = mdl.eDossier();
                //var NomFichier = mdl.eNomSansExt();

                var Barre = mdl.eSelect_RecupererObjet<Body2>(1);
                mdl.eEffacerSelection();

                var SM = mdl.SketchManager;

                WindowLog.Ecrire("Nom du corps : " + Barre.Name);

                var b = new Barre(Barre);

                //b.Analyser();

                WindowLog.Ecrire(b.Liste.Count);
                foreach (var Liste in b.Liste)
                {
                    mdl.eEffacerSelection();

                    SM.Insert3DSketch(false);
                    SM.AddToDB = true;
                    SM.DisplayWhenAdded = false;

                    foreach (var f in Liste)
                    {
                        mdl.eEffacerSelection();
                        f.eSelectEntite();
                        SM.SketchUseEdge3(true, false);
                    }

                    SM.DisplayWhenAdded = true;
                    SM.AddToDB = false;
                    SM.Insert3DSketch(true);

                    mdl.eEffacerSelection();
                }

                mdl.eEffacerSelection();

            }
            catch (Exception e) { this.LogMethode(new Object[] { e }); }

        }

        private List<Double> ListePercage(Body2 Barre)
        {
            Func<List<Edge>, Double> LgPercage = delegate (List<Edge> liste)
               {
                   double lg = 0;

                   foreach (var e in liste)
                   {
                       lg += e.eLgArrete();
                   }

                   return lg * 0.5;
               };

            var ListeListeArretes = new List<List<Edge>>();

            var ListeFoncCorps = Barre.eListeFonctions(null, false);

            if (ListeFoncCorps != null)
            {
                ListeFoncCorps.RemoveAt(0);

                var ListeFaces = new List<List<Edge>>();

                foreach (var Fonc in ListeFoncCorps)
                {
                    var ListeFoncFace = Fonc.eListeDesFaces();

                    foreach (var Face in ListeFoncFace)
                    {
                        var B = (Body2)Face.GetBody();
                        if (B.Name == Barre.Name)
                        {
                            var ListeBoucles = Face.eListeDesBoucles(l =>
                            {
                                if (l.IsOuter())
                                    return true;

                                return false;
                            });

                            // On ne recupère que les boucles exterieures
                            var ListeArrete = new List<Edge>();
                            foreach (var Boucle in ListeBoucles)
                            {
                                foreach (var Arrete in Boucle.GetEdges())
                                {
                                    ListeArrete.Add(Arrete);
                                }
                            }

                            ListeFaces.Add(ListeArrete);
                        }
                    }
                }

                while (ListeFaces.Count > 0)
                {
                    var ArreteFace1 = ListeFaces[0];
                    ListeListeArretes.Add(ArreteFace1);
                    ListeFaces.RemoveAt(0);

                    int index = 0;
                    while (index < ListeFaces.Count)
                    {
                        var ArreteFace2 = ListeFaces[index];
                        if (Union(ref ArreteFace1, ref ArreteFace2))
                        {
                            ListeFaces.RemoveAt(index);
                            index = -1;
                        }

                        index++;
                    }
                }

                WindowLog.Ecrire("Nb perçages : " + ListeListeArretes.Count);

                int i = 0;
                foreach (var liste in ListeListeArretes)
                {
                    WindowLog.Ecrire("Boucle " + i + " : " + liste.Count);
                    liste[0].eSelectEntite(true);
                }
            }

            var ListePercage = new List<Double>();

            foreach (var liste in ListeListeArretes)
            {
                ListePercage.Add(LgPercage(liste));
            }

            return ListePercage;
        }

        private Boolean Union(ref List<Edge> ListeArretes1, ref List<Edge> ListeArretes2)
        {
            Boolean Joindre = false;

            int i = 0;
            while (i < ListeArretes1.Count)
            {
                var Arrete1 = ListeArretes1[i];

                int j = 0;
                while (j < ListeArretes2.Count)
                {
                    var Arrete2 = ListeArretes2[j];

                    if (Arrete1.eIsSame(Arrete2))
                    {
                        Joindre = true;

                        ListeArretes1.RemoveAt(i);
                        ListeArretes2.RemoveAt(j);
                        i--;
                        break;
                    }

                    j++;
                }
                i++;
            }

            if (Joindre)
            {
                ListeArretes1.AddRange(ListeArretes2);
                return true;
            }

            return false;
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
