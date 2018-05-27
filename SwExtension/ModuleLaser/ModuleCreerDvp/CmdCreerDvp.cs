using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;

namespace ModuleLaser.ModuleCreerDvp
{
    public class CmdCreerDvp : Cmd
    {
        public ModelDoc2 MdlBase = null;
        public List<String> ListeMateriaux = new List<String>();
        public String ForcerMateriau = null;
        public int Quantite = 1;
        public Boolean AfficherLignePliage = false;
        public Boolean AfficherNotePliage = false;
        public Boolean InscrireNomTole = false;
        public Boolean OrienterDvp = false;
        public eOrientation OrientationDvp = eOrientation.Portrait;
        public Boolean FermerPlan = false;
        public Boolean ConvertirEsquisse = false;
        public Boolean MajPlans = false;
        public Boolean ComposantsExterne = false;
        public String RefFichier = "";
        public int TailleInscription = 5;

        private Dictionary<String, DrawingDoc> DicDessins = new Dictionary<string, DrawingDoc>();

        private List<String> DicErreur = new List<String>();

        private String DossierDVP = "";

        private String ResumeLog = "";

        protected override void Command()
        {
            try
            {
                if (ConvertirEsquisse)
                {
                    WindowLog.Ecrire("Attention !!!");
                    WindowLog.Ecrire("Les dvp seront convertis en esquisse");
                    WindowLog.SautDeLigne();
                }

                CreerDossierDVP();

                eTypeCorps Filtre = eTypeCorps.Tole;
                HashSet<String> HashMateriaux = new HashSet<string>(ListeMateriaux);

                var dic = MdlBase.DenombrerDossiers(ComposantsExterne, HashMateriaux, Filtre);

                int MdlPct = 0;
                foreach (var mdl in dic.Keys)
                {
                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                    WindowLog.SautDeLigne();
                    WindowLog.EcrireF("[{1}/{2}] {0}", mdl.eNomSansExt(), ++MdlPct, dic.Count);
                    

                    int CfgPct = 0;
                    foreach (var NomConfigPliee in dic[mdl].Keys)
                    {
                        WindowLog.SautDeLigne();
                        WindowLog.EcrireF("  [{1}/{2}] Config : \"{0}\"", NomConfigPliee, ++CfgPct, dic[mdl].Count);
                        mdl.ShowConfiguration2(NomConfigPliee);
                        mdl.EditRebuild3();
                        PartDoc Piece = mdl.ePartDoc();

                        var ListeDossier = dic[mdl][NomConfigPliee];
                        int DrPct = 0;
                        foreach (var t in ListeDossier)
                        {
                            var IdDossier = t.Key;
                            var QuantiteTole = t.Value * Quantite;

                            Feature fDossier = Piece.FeatureById(IdDossier);
                            BodyFolder dossier = fDossier.GetSpecificFeature2();

                            var RefDossier = dossier.eProp(CONSTANTES.REF_DOSSIER);

                            Body2 Tole = dossier.eCorpsDeTolerie();

                            WindowLog.SautDeLigne();
                            WindowLog.EcrireF("    - [{1}/{2}] Dossier : \"{0}\" x{3}", RefDossier, ++DrPct, ListeDossier.Count, QuantiteTole);

                            String Materiau = Tole.eGetMateriauCorpsOuPiece(Piece, NomConfigPliee);

                            Materiau = ForcerMateriau.IsRefAndNotEmpty(Materiau);

                            Double Epaisseur = Tole.eEpaisseur();
                            String NomConfigDepliee = Sw.eNomConfigDepliee(NomConfigPliee, RefDossier);

                            WindowLog.EcrireF("      Ep {0} / Materiau {1}", Epaisseur, Materiau);
                            WindowLog.EcrireF("          Config {0}", NomConfigDepliee);

                            if (ConvertirEsquisse)
                            {
                                if (!mdl.eListeNomConfiguration().Contains(NomConfigDepliee))
                                    mdl.CreerConfigDepliee(NomConfigDepliee, NomConfigPliee);

                                WindowLog.EcrireF("          Configuration crée : {0}", NomConfigDepliee);
                                mdl.ShowConfiguration2(NomConfigDepliee);
                                Tole.DeplierTole(mdl, NomConfigDepliee);
                            }
                            else if (!mdl.ShowConfiguration2(NomConfigDepliee))
                            {
                                DicErreur.Add(mdl.eNomSansExt() + " -> cfg : " + NomConfigPliee + " - No : " + RefDossier + " = " + NomConfigDepliee);
                                WindowLog.EcrireF("La configuration n'éxiste pas");
                                continue;
                            }

                            mdl.EditRebuild3();

                            DrawingDoc dessin = CreerPlan(Materiau, Epaisseur);
                            dessin.eModelDoc2().eActiver();
                            Sheet Feuille = dessin.eFeuilleActive();

                            View v = CreerVueToleDvp(dessin, Feuille, Piece, NomConfigDepliee, RefDossier, Materiau, QuantiteTole, Epaisseur);

                            if (ConvertirEsquisse)
                            {
                                mdl.ShowConfiguration2(NomConfigPliee);
                                mdl.EditRebuild3();
                                mdl.DeleteConfiguration2(NomConfigDepliee);
                            }
                        }

                        NouvelleLigne = true;
                    }

                    if (mdl.GetPathName() != MdlBase.GetPathName())
                        App.Sw.CloseDoc(mdl.GetPathName());
                }

                foreach (DrawingDoc dessin in DicDessins.Values)
                {
                    int Errors = 0, Warnings = 0;
                    dessin.eModelDoc2().eActiver();
                    dessin.eFeuilleActive().eAjusterAutourDesVues();
                    dessin.eModelDoc2().ViewZoomtofit2();
                    dessin.eModelDoc2().Save3((int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced + (int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref Errors, ref Warnings);
                }

                if (FermerPlan)
                {
                    foreach (ModelDoc2 dessin in DicDessins.Values)
                        App.Sw.CloseDoc(dessin.GetPathName());
                }

                if (DicErreur.Count > 0)
                {
                    WindowLog.SautDeLigne();

                    WindowLog.Ecrire("Liste des erreurs :");
                    foreach (var item in DicErreur)
                        WindowLog.Ecrire(" - " + item);

                    WindowLog.SautDeLigne();
                }
                else
                {
                    WindowLog.Ecrire("Pas d'erreur");
                }

                File.WriteAllText(Path.Combine(DossierDVP, "Log_CreerDvp.txt"), WindowLog.Resume);
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private String cmdRefPiece(ModelDoc2 mdl, String configPliee, String noDossier)
        {
            return String.Format("{0}-{1}-{2}", mdl.eNomSansExt(), configPliee, noDossier);
        }

        private Boolean NouvelleLigne = false;

        private Dictionary<String, eZone> DicPoint = new Dictionary<string, eZone>();

        private void PositionnerVue(Sheet feuille, View vue, GeomVue g)
        {
            Double JeuEntreVue = 0.05;
            eZone z;
            if (DicPoint.ContainsKey(feuille.GetName()))
            {
                z = DicPoint[feuille.GetName()];
                z.PointMax.X += JeuEntreVue;
            }
            else
            {
                z = new eZone();
                Double Ymin = JeuEntreVue;

                if (feuille.eListeDesVues().Count > 0)
                {
                    var e = feuille.eEnveloppeDesVues();
                    Ymin += e.PointMax.Y;
                }

                z.PointMin.X = JeuEntreVue; z.PointMin.Y = Ymin;
                z.PointMax.X = z.PointMin.X; z.PointMax.Y = z.PointMin.Y;
                DicPoint.Add(feuille.GetName(), z);
            }

            if (z.PointMax.X > 10)
                NouvelleLigne = true;

            if (NouvelleLigne)
            {
                z.PointMin.Y = z.PointMax.Y + JeuEntreVue;
                z.PointMax.X = z.PointMin.X;
                z.PointMax.Y = z.PointMin.Y;

                NouvelleLigne = false;
            }

            Double[] P = new Double[2];
            P[0] = z.PointMax.X + (g.ptCentreVueX - g.ptMinX);
            P[1] = z.PointMin.Y + (g.ptCentreVueY - g.ptMinY);

            vue.Position = P;

            z.PointMax.X = z.PointMax.X + g.Lg + JeuEntreVue;
            z.PointMax.Y = Math.Max(z.PointMax.Y, z.PointMin.Y + g.Ht);
        }

        private void CreerDossierDVP()
        {
            DossierDVP = System.IO.Path.Combine(MdlBase.eDossier(), CONSTANTES.DOSSIER_DVP);
            if (!System.IO.Directory.Exists(DossierDVP))
                System.IO.Directory.CreateDirectory(DossierDVP);
        }

        private DrawingDoc CreerPlan(String materiau, Double epaisseur)
        {
            String Fichier = String.Format("{0}{1} - Ep{2}",
                                            String.IsNullOrWhiteSpace(RefFichier) ? "" : RefFichier + " - ",
                                            materiau.eGetSafeFilename("-"),
                                            epaisseur.ToString().Replace('.', ',')
                                            ).Trim();

            if (DicDessins.ContainsKey(Fichier))
                return DicDessins[Fichier];

            if (MajPlans)
            {
                var di = new DirectoryInfo(DossierDVP);
                var NomFichierExt = Fichier + eTypeDoc.Dessin.GetEnumInfo<ExtFichier>();
                var r = di.GetFiles(NomFichierExt, SearchOption.TopDirectoryOnly);
                if (r.Length > 0)
                {
                    int Erreur = 0, Warning = 0;
                    DrawingDoc dessin = App.Sw.OpenDoc6(Path.Combine(DossierDVP, NomFichierExt),
                        (int)swDocumentTypes_e.swDocDRAWING,
                        (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                        "",
                        ref Erreur,
                        ref Warning).eDrawingDoc();
                    DicDessins.Add(Fichier, dessin);

                    return dessin;
                }
            }

            DrawingDoc Dessin = Sw.eCreerDocument(DossierDVP, Fichier, eTypeDoc.Dessin, CONSTANTES.MODELE_DE_DESSIN_LASER).eDrawingDoc();

            Dessin.eFeuilleActive().SetName(Fichier);
            DicDessins.Add(Fichier, Dessin);

            ModelDoc2 mdl = Dessin.eModelDoc2();

            LayerMgr LM = mdl.GetLayerManager();
            LM.AddLayer("GRAVURE", "", 1227327, (int)swLineStyles_e.swLineCONTINUOUS, (int)swLineWeights_e.swLW_LAYER);
            LM.AddLayer("QUANTITE", "", 1227327, (int)swLineStyles_e.swLineCONTINUOUS, (int)swLineWeights_e.swLW_LAYER);

            ModelDocExtension ext = mdl.Extension;

            ext.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swFlatPatternOpt_ShowFixedFace, 0, false);
            ext.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swShowSheetMetalBendNotes, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, AfficherNotePliage);

            TextFormat tf = ext.GetUserPreferenceTextFormat(((int)(swUserPreferenceTextFormat_e.swDetailingAnnotationTextFormat)), 0);
            tf.CharHeight = TailleInscription / 1000.0;
            ext.SetUserPreferenceTextFormat((int)swUserPreferenceTextFormat_e.swDetailingAnnotationTextFormat, 0, tf);

            return Dessin;
        }

        public View CreerVueToleDvp(DrawingDoc dessin, Sheet feuille, PartDoc piece, String configDepliee, String Ref, String materiau, int quantite, Double epaisseur)
        {
            // Si des corps autre que la tole dépliée sont encore visible dans la config, on les cache et on recontruit tout
            foreach (Body2 pCorps in piece.eListeCorps(false))
            {
                if ((pCorps.Name == CONSTANTES.NOM_CORPS_DEPLIEE))
                    pCorps.eVisible(true);
                else
                    pCorps.eVisible(false);
            }

            var NomVue = piece.eModelDoc2().eNomSansExt() + " - " + configDepliee;

            if (MajPlans)
            {
                foreach (var vue in feuille.eListeDesVues())
                {
                    if (vue.GetName2() == NomVue)
                    {
                        vue.eSelectionner(dessin);
                        dessin.eModelDoc2().Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Advanced);
                        break;
                    }
                }
            }

            dessin.eModelDoc2().eEffacerSelection();

            View Vue = dessin.CreateFlatPatternViewFromModelView3(piece.eModelDoc2().GetPathName(), configDepliee, 0, 0, 0, !AfficherLignePliage, false);
            Vue.ScaleDecimal = 1;

            if (Vue.IsRef())
            {
                Vue.SetName2(NomVue);

                GeomVue g = AppliquerOptionsVue(dessin, piece, Vue);

                Vue.eSelectionner(dessin);

                {
                    Note Note = dessin.eModelDoc2().InsertNote(String.Format("{0}× {1} [{2}] (Ep{3})", quantite, Ref, materiau, epaisseur));
                    Note.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);

                    Annotation Annotation = Note.GetAnnotation();

                    TextFormat swTextFormat = Annotation.GetTextFormat(0);
                    // Hauteur du texte en fonction des dimensions du dvp, 2.5% de la dimension max du dvp
                    swTextFormat.CharHeight = Math.Max(0.005, Math.Floor(Math.Max(g.Ht, g.Lg) * 0.025 * 1000) * 0.001);
                    Annotation.SetTextFormat(0, false, swTextFormat);

                    Annotation.Layer = "QUANTITE";
                    Annotation.SetLeader3((int)swLeaderStyle_e.swNO_LEADER, (int)swLeaderSide_e.swLS_SMART, true, true, false, false);
                    // Décalage du texte en fonction des dimensions du dvp, 0.1% de la dimension max du dvp
                    Annotation.SetPosition(g.ptCentreRectangleX, g.ptMinY - Math.Max(0.005, Math.Floor(Math.Max(g.Ht, g.Lg) * 0.001 * 1000) * 0.001), 0);
                    g.Agrandir(Note);
                }

                if (InscrireNomTole)
                {
                    Note Note = dessin.eModelDoc2().InsertNote(Ref);
                    Note.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);

                    Annotation Annotation = Note.GetAnnotation();
                    Annotation.Layer = "GRAVURE";
                    Annotation.SetLeader3((int)swLeaderStyle_e.swNO_LEADER, (int)swLeaderSide_e.swLS_SMART, true, true, false, false);
                }

                if (ConvertirEsquisse)
                {
                    String titre = Vue.GetName2();
                    Vue.ReplaceViewWithSketch();
                    Vue = feuille.eListeDesVues().Last();
                    Vue.SetName2(titre);
                    g = new GeomVue(Vue);
                }

                PositionnerVue(feuille, Vue, g);

                return Vue;
            }

            return null;
        }

        private GeomVue AppliquerOptionsVue(DrawingDoc dessin, PartDoc piece, View vue)
        {
            vue.ShowSheetMetalBendNotes = AfficherNotePliage;

            MathUtility SwMath = App.Sw.GetMathUtility();

            ModelDoc2 mdl = dessin.eModelDoc2();

            Body2 c = piece.eChercherCorps(CONSTANTES.NOM_CORPS_DEPLIEE, true);

            GeomVue g = null;

            c.eFonctionEtatDepliee().eParcourirSousFonction(
                f =>
                {
                    if (f.Name.StartsWith(CONSTANTES.LIGNES_DE_PLIAGE))
                    {
                        String NomSelection = f.Name + "@" + vue.RootDrawingComponent.Name + "@" + vue.Name;
                        mdl.Extension.SelectByID2(NomSelection, "SKETCH", 0, 0, 0, false, 0, null, 0);
                        if (AfficherLignePliage)
                            mdl.UnblankSketch();
                        else
                            mdl.BlankSketch();

                        mdl.eEffacerSelection();
                    }
                    else if (f.Name.StartsWith(CONSTANTES.CUBE_DE_VISUALISATION))
                    {
                        try
                        {
                            f.eModifierEtat(swFeatureSuppressionAction_e.swUnSuppressFeature);

                            Sketch Esquisse = f.GetSpecificFeature2();

                            if (OrienterDvp)
                            {
                                Double Angle = AngleCubeDeVisualisation(vue, Esquisse);

                                // On oriente la vue
                                switch (OrientationDvp)
                                {
                                    case eOrientation.Portrait:
                                        if (Math.Abs(Angle) != MathX.Rad90D)
                                        {
                                            Double a = MathX.Rad90D - Math.Abs(Angle);
                                            vue.Angle = (Math.Sign(Angle) == 0 ? 1 : Math.Sign(Angle)) * a;
                                        }
                                        break;
                                    case eOrientation.Paysage:
                                        if (Math.Abs(Angle) != 0 || Math.Abs(Angle) != MathX.Rad180D)
                                        {
                                            Double a = MathX.Rad90D - Math.Abs(Angle);
                                            vue.Angle = ((Math.Sign(Angle) == 0 ? 1 : Math.Sign(Angle)) * a) - MathX.Rad90D;
                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }

                            g = new GeomVue(vue, Esquisse);
                        }
                        catch
                        {
                            WindowLog.EcrireF("Impossible d'orienter la vue : {0}", vue.Name);
                        }
                        return true;
                    }
                    return false;
                }
                );

            return g;
        }

        private Double AngleCubeDeVisualisation(View vue, Sketch esquisse)
        {
            MathUtility SwMath = App.Sw.GetMathUtility();

            List<Point> LstPt = new List<Point>();
            foreach (SketchPoint s in esquisse.GetSketchPoints2())
            {
                MathPoint point = SwMath.CreatePoint(new Double[3] { s.X, s.Y, s.Z });
                MathTransform SketchXform = esquisse.ModelToSketchTransform;
                SketchXform = SketchXform.Inverse();
                point = point.MultiplyTransform(SketchXform);
                MathTransform ViewXform = vue.ModelToViewTransform;
                point = point.MultiplyTransform(ViewXform);
                Point swViewStartPt = new Point(point);
                LstPt.Add(swViewStartPt);
            }

            // On recherche le point le point le plus à droite puis le plus haut
            LstPt.Sort(new PointComparer(ListSortDirection.Descending, p => p.X));
            LstPt.Sort(new PointComparer(ListSortDirection.Descending, p => p.Y));

            // On le supprime
            LstPt.RemoveAt(0);

            // On recherche le point le point le plus à gauche puis le plus bas 
            LstPt.Sort(new PointComparer(ListSortDirection.Ascending, p => p.X));
            LstPt.Sort(new PointComparer(ListSortDirection.Ascending, p => p.Y));


            // C'est le point de rotation
            Point pt1 = LstPt[0];

            // On recherche le plus loin
            Point pt2;
            Double d1 = pt1.Distance(LstPt[1]);
            Double d2 = pt1.Distance(LstPt[2]);

            if (d1 > d2)
            {
                pt2 = LstPt[1];
            }
            // En cas d'égalité, on renvoi le point le plus à gauche
            else if (d1 == d2)
            {
                pt2 = (LstPt[1].X < LstPt[2].X) ? LstPt[1] : LstPt[2];
            }
            else
                pt2 = LstPt[2];

            Vecteur v = new Vecteur(pt1, pt2);

            return Math.Atan2(v.Y, v.X);
        }

        private class GeomVue
        {
            public Double ptCentreVueX = 0;
            public Double ptCentreVueY = 0;
            public Double ptMinX = 0;
            public Double ptMinY = 0;

            public Double ptMaxX = 0;
            public Double ptMaxY = 0;

            public Double ptCentreRectangleX = 0;
            public Double ptCentreRectangleY = 0;

            public GeomVue(View vue, Sketch esquisse)
            {
                MathUtility SwMath = App.Sw.GetMathUtility();

                Point ptCentreVue = new Point(vue.Position);
                Point ptMin = new Point(Double.PositiveInfinity, Double.PositiveInfinity, 0);
                Point ptMax = new Point(Double.NegativeInfinity, Double.NegativeInfinity, 0);

                foreach (SketchPoint s in esquisse.GetSketchPoints2())
                {
                    MathPoint swStartPoint = SwMath.CreatePoint(new Double[3] { s.X, s.Y, s.Z });
                    MathTransform SketchXform = esquisse.ModelToSketchTransform;
                    SketchXform = SketchXform.Inverse();
                    swStartPoint = swStartPoint.MultiplyTransform(SketchXform);
                    MathTransform ViewXform = vue.ModelToViewTransform;
                    swStartPoint = swStartPoint.MultiplyTransform(ViewXform);
                    Point swViewStartPt = new Point(swStartPoint);
                    ptMin.Min(swViewStartPt);
                    ptMax.Max(swViewStartPt);
                }

                ptMinX = ptMin.X;
                ptMinY = ptMin.Y;
                ptMaxX = ptMax.X;
                ptMaxY = ptMax.Y;
                MajCentreRectangle();
                ptCentreVueX = ptCentreVue.X;
                ptCentreVueY = ptCentreVue.Y;
            }

            public GeomVue(View vue)
            {
                MathUtility SwMath = App.Sw.GetMathUtility();

                Point ptCentreVue = new Point(vue.Position);
                Double[] pArr = vue.GetOutline();
                ptMinX = pArr[0];
                ptMinY = pArr[1];
                ptMaxX = pArr[2];
                ptMaxY = pArr[3];

                foreach (Note n in vue.GetNotes())
                    Agrandir(n);

                MajCentreRectangle();
                ptCentreVueX = ptCentreVue.X;
                ptCentreVueY = ptCentreVue.Y;
            }

            private void MajCentreRectangle()
            {
                ptCentreRectangleX = (ptMinX + ptMaxX) * 0.5;
                ptCentreRectangleY = (ptMinY + ptMaxY) * 0.5;
            }

            public void Agrandir(Point p)
            {
                ptMinX = Math.Min(ptMinX, p.X);
                ptMinY = Math.Min(ptMinY, p.Y);

                ptMaxX = Math.Max(ptMaxX, p.X);
                ptMaxY = Math.Max(ptMaxY, p.Y);

                MajCentreRectangle();
            }

            public void Agrandir(Note p)
            {
                Double[] dimNote = (Double[])p.GetExtent();
                Agrandir(new Point(dimNote[0], dimNote[1], dimNote[2]));
                Agrandir(new Point(dimNote[3], dimNote[4], dimNote[5]));

                MajCentreRectangle();
            }

            public Double Lg
            {
                get { return ptMaxX - ptMinX; }
            }

            public Double Ht
            {
                get { return ptMaxY - ptMinY; }
            }
        }
    }
}


