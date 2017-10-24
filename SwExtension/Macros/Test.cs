using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using TriangleNet;
using TriangleNet.Geometry;
using TriangleNet.Meshing;
using TriangleNet.Voronoi;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Test"),
        ModuleNom("Test")]
    public class Test : BoutonBase
    {

        private List<String> CalquesBase = new List<string>() {"Annotations", "Cotations", "Tables", "Vue", "Construction", "Bordure", "Pliage"};

        protected override void Command()
        {
            try
            {
                ModelDoc2 MdlBase = App.ModelDoc2;

                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swWeldmentEnableAutomaticCutList, 0, true);
                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swWeldmentEnableAutomaticUpdate, 0, false);
                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisableDerivedConfigurations, 0, false);
                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swWeldmentRenameCutlistDescriptionPropertyValue, 0, true);
                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swWeldmentCollectIdenticalBodies, 0, true);
                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSheetMetalBodiesDescriptionUseDefault, 0, false);
                MdlBase.Extension.SetUserPreferenceString((int)swUserPreferenceStringValue_e.swSheetMetalDescription, 0, "Tôle");
                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swFlatPatternOpt_SimplifyBends, 0, false);
                MdlBase.Extension.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swFlatPatternOpt_CornerTreatment, 0, false);

                if (MdlBase.TypeDoc() == eTypeDoc.Dessin)
                {

                    LayerMgr LM = MdlBase.GetLayerManager();

                    String[] ListeCalques = LM.GetLayerList();

                    WindowLog.Ecrire(MdlBase.GetPathName());

                    foreach (var Calque in ListeCalques)
                    {
                        if (!CalquesBase.Contains(Calque))
                        {
                            WindowLog.Ecrire(Calque + " : " + LM.DeleteLayer(Calque));
                        }
                        else
                        {
                            WindowLog.Ecrire(Calque);
                        }
                    }

                    String cheminDossier = App.Sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsSheetFormat) + "\\";
                    String cheminFondPlan = MdlBase.eDrawingDoc().eFeuilleActive().eGetGabaritDeFeuille();

                    String nomFondPlan = cheminFondPlan.Replace(cheminDossier, "");

                    WindowLog.Ecrire(nomFondPlan);

                    if (nomFondPlan.ToLower().StartsWith("archi"))
                        MdlBase.Extension.LoadDraftingStandard("E:\\Mes documents\\SolidWorks\\2018\\Norme dessin\\Norme Archi.sldstd");
                    else
                        MdlBase.Extension.LoadDraftingStandard("E:\\Mes documents\\SolidWorks\\2018\\Norme dessin\\Norme Fab.sldstd");

                    MdlBase.eDrawingDoc().eFeuilleActive().SaveFormat(cheminFondPlan);
                }
                else
                {
                    int lErrors = 0;
                    int lWarnings = 0;

                    MdlBase.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref lErrors, ref lWarnings);
                }

            }
            catch (Exception e) { this.LogMethode(new Object[] { e }); }

        }

        //private Boolean Decompter(Component2 cp)
        //{
        //    try
        //    {
        //        if (!cp.IsHidden(true))
        //        {
        //            foreach (var corps in cp.eListeCorps())
        //            {
        //                foreach (var face in corps.eListeDesFaces())
        //                {
        //                    Surface S = face.GetSurface();
        //                    if (S.IsCylinder() && (face.GetLoopCount() > 1))
        //                    {
        //                        Double[] ListeParam = (Double[])S.CylinderParams;
        //                        Double Diam = Math.Round(ListeParam[6] * 2.0 * 1000, 2);

        //                        DicQte.AddIfNotExistOrPlus(Diam);
        //                    }
        //                }
        //            }
        //        }

        //    }
        //    catch (Exception e) { this.LogMethode(new Object[] { e }); }

        //    return false;
        //}
    }

    //public class Test : BoutonBase
    //{
    //    private MathTransform xForm = null;
    //    private MathUtility Mu = null;
    //    private Sketch Esquisse = null;
    //    private Double FacteurDecal = 0.8;

    //    protected override void Command()
    //    {

    //        ModelDoc2 MdlBase = App.ModelDoc2;

    //        String[] tabs = MdlBase.Extension.GetCommandTabs();

    //        foreach (var tab in tabs)
    //        {
    //            WindowLog.Ecrire(tab);
    //            //WindowLog.Ecrire(tab.Name);
    //            //WindowLog.Ecrire("   " + tab.GetCommandTabBoxCount());
    //            //WindowLog.SautDeLigne();
    //        }

    //        WindowLog.Ecrire(MdlBase.Extension.ActiveCommandTab);

    //    }

    //    private void CreerVoronoi()
    //    {
    //        ModelDoc2 MdlBase = App.ModelDoc2;

    //        Feature Fonction = MdlBase.eSelect_RecupererObjet<Feature>(1, -1);

    //        MdlBase.eEffacerSelection();

    //        Esquisse = Fonction.GetSpecificFeature2();
    //        xForm = Esquisse.ModelToSketchTransform.Inverse();
    //        Mu = App.Sw.GetMathUtility();

    //        String fact = "0.8";

    //        if (Interaction.InputBox("Facteur de décalage", "F :", ref fact) == DialogResult.OK)
    //        {
    //            if (!String.IsNullOrWhiteSpace(fact))
    //            {
    //                Double r = fact.eToDouble();
    //                if (r != 0)
    //                    FacteurDecal = r;
    //            }
    //        }

    //        var polygon = new Polygon();

    //        if (Esquisse.GetSketchRegionCount() == 0) return;

    //        // On ajoute les contours
    //        SketchRegion region = Esquisse.GetSketchRegions()[0];
    //        Loop2 loop = region.GetFirstLoop();

    //        int i = 0;
    //        while (loop != null)
    //        {
    //            var ListPt = new List<TriangleNet.Geometry.Vertex>();
    //            foreach (SolidWorks.Interop.sldworks.Vertex v in loop.GetVertices())
    //            {
    //                double[] pt = (double[])v.GetPoint();
    //                ListPt.Add(new TriangleNet.Geometry.Vertex(pt[0], pt[1]));
    //            }

    //            polygon.Add(new Contour(ListPt, i++), !loop.IsOuter());

    //            loop = loop.GetNext();
    //        }


    //        foreach (SketchPoint pt in Esquisse.GetSketchPoints2())
    //            polygon.Add(new TriangleNet.Geometry.Vertex(pt.X, pt.Y));

    //        var contraintes = new ConstraintOptions();
    //        contraintes.ConformingDelaunay = true;
    //        contraintes.SegmentSplitting = 0;

    //        //var Qualite = new QualityOptions();
    //        //Qualite.SteinerPoints = 0;
    //        Mesh mesh = (Mesh)polygon.Triangulate(contraintes);

    //        var voronoi = new BoundedVoronoi(mesh);

    //        var SM = MdlBase.SketchManager;

    //        SelectionnerSupportEsquisse();
    //        SM.InsertSketch(false);
    //        SM.AddToDB = true;
    //        SM.DisplayWhenAdded = false;

    //        List<TriangleNet.Geometry.Vertex> lstVertex = new List<TriangleNet.Geometry.Vertex>();
    //        foreach (var v in mesh.Vertices)
    //        {
    //            lstVertex.Add(v);
    //        }

    //        foreach (var edge in mesh.Edges)
    //        {
    //            TriangleNet.Geometry.Point p1 = TransformPt(lstVertex[edge.P0]);
    //            TriangleNet.Geometry.Point p2 = TransformPt(lstVertex[edge.P1]);
    //            SM.CreateLine(p1.X, p1.Y, 0, p2.X, p2.Y, 0);
    //        }

    //        SM.DisplayWhenAdded = true;
    //        SM.AddToDB = false;
    //        SM.InsertSketch(true);
    //        MdlBase.eEffacerSelection();

    //        SelectionnerSupportEsquisse();
    //        SM.InsertSketch(false);
    //        SM.AddToDB = true;
    //        SM.DisplayWhenAdded = false;

    //        foreach (var f in voronoi.Faces)
    //        {
    //            try
    //            {
    //                TriangleNet.Geometry.Point centre = f.generator;

    //                foreach (var e in f.EnumerateEdges())
    //                {
    //                    var pt1 = TransformPt(Echelle(centre, e.Origin, FacteurDecal));
    //                    var pt2 = TransformPt(Echelle(centre, e.Twin.Origin, FacteurDecal));
    //                    SM.CreateLine(pt1.X, pt1.Y, 0, pt2.X, pt2.Y, 0);
    //                }
    //            }
    //            catch (Exception errDesc)
    //            { this.LogMethode(new Object[] { errDesc }); }
    //        }

    //        SM.DisplayWhenAdded = true;
    //        SM.AddToDB = false;
    //        SM.InsertSketch(true);
    //    }

    //    private void SelectionnerSupportEsquisse()
    //    {
    //        int TypeRef = 0;
    //        Object obj = Esquisse.GetReferenceEntity(ref TypeRef);
    //        if (TypeRef == (int)swSelectType_e.swSelFACES)
    //        {
    //            Entity ent = (Entity)obj;
    //            ent.eSelectEntite();
    //        }
    //        else
    //        {
    //            RefPlane plan = (RefPlane)obj;
    //            Feature f = (Feature)obj;
    //            f.eSelect();
    //        }
    //    }

    //    private TriangleNet.Geometry.Point TransformPt(TriangleNet.Geometry.Point Pt)
    //    {
    //        double[] arr = new double[3];
    //        arr[0] = Pt.X;
    //        arr[1] = Pt.Y;
    //        arr[2] = 0;
    //        MathPoint swPt = (MathPoint)Mu.CreatePoint(arr);

    //        swPt = swPt.MultiplyTransform(xForm);
    //        arr = swPt.ArrayData;

    //        return new TriangleNet.Geometry.Point(arr[0], arr[1]);
    //    }

    //    private TriangleNet.Geometry.Point Echelle(TriangleNet.Geometry.Point centre, TriangleNet.Geometry.Point point, double f)
    //    {
    //        Vecteur v = new Vecteur(point.X - centre.X, point.Y - centre.Y, 0);
    //        v.Multiplier(f);
    //        double x = centre.X + v.X;
    //        double y = centre.Y + v.Y;

    //        return new TriangleNet.Geometry.Point(x, y);
    //    }
    //}
}
