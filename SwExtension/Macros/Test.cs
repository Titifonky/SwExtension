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

            public Plan PlanSection;
            public List<Face2> ListeFaceSection = new List<Face2>();

            public List<Face2> ListeFaceExtremite1;

            public List<Face2> ListeFaceExtremite2;

            public Face2 FaceBase;

            public Barre(Body2 corps)
            {
                Corps = corps;

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

                Boolean Extrusion = false;

                List<FaceExt> ListeFaceExt = new List<FaceExt>();

                // Tri des faces pour retrouver celles issues de la même
                foreach (var Face in ListeFaces)
                {
                    var faceExt = new FaceExt(Face);

                    if (faceExt.Type != eTypeFace.Plan)
                        Extrusion = true;

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

                //var SM = Sw.eModeleActif().SketchManager;

                //foreach (var f in ListeFaceExt)
                //{
                //    if (f.Type == eTypeFace.Plan)
                //    {
                //        SM.Insert3DSketch(false);
                //        SM.AddToDB = true;
                //        SM.DisplayWhenAdded = false;

                //        var O = f.Origine;
                //        var N = f.Normale;
                //        SM.CreatePoint(O.X, O.Y, O.Z);
                //        SM.CreateLine(O.X, O.Y, O.Z, O.X + (N.X * 0.005), O.Y + (N.Y * 0.005), O.Z + (N.Z * 0.005));

                //        SM.DisplayWhenAdded = true;
                //        SM.AddToDB = false;
                //        SM.Insert3DSketch(true);
                //    }
                //}

                var ListeTest = new List<FaceExt>();
                foreach (var f in ListeFaceExt)
                {
                    if(f.Type == eTypeFace.Plan)
                        ListeTest.Add(f);
                }

                FaceExt fStart;

                int milieu = (ListeTest.Count / 2) - 1;
                fStart = ListeTest[milieu];
                ListeTest.RemoveAt(milieu);
                FaceBase = fStart.SwFace;

                var DicPlan = new Dictionary<Plan, List<FaceExt>>();

                // On recherche les plans
                foreach (var f in ListeTest)
                {
                    var test = Orientation(f, fStart);

                    if (test == eOrientation.Coplanaire || test == eOrientation.MemeOrigine)
                    {
                        var vF = (new Vecteur(fStart.Origine, f.Origine)).Compose(f.Normale);
                        var v = fStart.Normale.Vectoriel(vF);
                        var plan = new Plan(fStart.Origine, v);

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
                            DicPlan.Add(plan, new List<FaceExt>() { f });
                    }
                }

                // On regarde si des points se retrouve sur des plans
                foreach (var f in ListeTest)
                {
                    var test = Orientation(f, fStart);
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

                // On recherche le plan qui contient le plus de face
                Plan Pmax = new Plan();
                var ListeMax = new List<FaceExt>();

                WindowLog.Ecrire(DicPlan.Count);
                foreach (var p in DicPlan.Keys)
                {
                    var l = DicPlan[p];
                    if (ListeMax.Count < l.Count)
                    {
                        Pmax = p;
                        ListeMax = l;
                    }
                }

                

                PlanSection = Pmax;

                ListeMax.Add(fStart);
                foreach (var fe in ListeMax)
                {
                    ListeFaceSection.AddRange(fe.ListeSwFace);
                    ListeFaceExt.Remove(fe);
                }
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

            public class FaceExt
            {
                public Face2 SwFace = null;
                private Surface Surface = null;

                public Point Origine;
                public Vecteur Normale;
                public Vecteur Direction;
                public eTypeFace Type = eTypeFace.Inconnu;

                public List<Face2> ListeSwFace = new List<Face2>();

                public FaceExt(Face2 swface)
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

                public Boolean FaceExtIdentique(FaceExt fe, Double arrondi = 1E-10)
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
                            if (!Direction.EstColineaire(fe.Direction, arrondi))
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

                    if(Surface.IsPlane())
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
                    Boolean Reverse = SwFace.FaceInSurfaceSense();

                    if (Surface.IsCylinder())
                    {
                        Double[] Param = Surface.CylinderParams;

                        Origine = new Point(Param[0], Param[1], Param[2]);
                        Direction = new Vecteur(Param[3], Param[4], Param[5]);
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

                        if(C.GetEndParams(out StartParam, out EndParam, out IsClosed, out IsPeriodic))
                        {
                            Double[] Eval = C.Evaluate(StartParam);

                            Origine = new Point(Eval[0], Eval[1], Eval[2]);
                        }
                    }
                }
            }

            public eOrientation Orientation(FaceExt f1, FaceExt f2)
            {
                var val = eOrientation.Indefini;
                if (f1.Type == eTypeFace.Plan && f2.Type == eTypeFace.Plan)
                {
                    val = Orientation(f1.Origine, f1.Normale, f2.Origine, f2.Normale);
                }

                return val;
            }

            public eOrientation Orientation(Point p1, Vecteur v1, Point p2, Vecteur v2)
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

                WindowLog.Ecrire(b.ListeFaceSection.Count);

                foreach (var f in b.ListeFaceSection)
                {
                    f.eSelectEntite(true);
                }

                //b.FaceBase.eSelectEntite();

                //foreach (var Liste in b.Liste)
                //{
                //    mdl.eEffacerSelection();

                //    SM.Insert3DSketch(false);
                //    SM.AddToDB = true;
                //    SM.DisplayWhenAdded = false;

                //    foreach (var f in Liste)
                //    {
                //        mdl.eEffacerSelection();
                //        f.eSelectEntite();
                //        SM.SketchUseEdge3(true, false);
                //    }

                //    SM.DisplayWhenAdded = true;
                //    SM.AddToDB = false;
                //    SM.Insert3DSketch(true);

                //    mdl.eEffacerSelection();
                //}

                //mdl.eEffacerSelection();

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
