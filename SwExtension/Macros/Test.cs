using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Test"),
        ModuleNom("Test")]

    public class Test : BoutonBase
    {
        public static int index = -1;

        protected override void Command()
        {
            try
            {
                ModelDoc2 MdlBase = App.ModelDoc2;

                var f = MdlBase.eSelect_RecupererObjet<Face2>();

                MdlBase.eEffacerSelection();

                Body2 b = f.GetBody();

                String val = "Test";

                if(index == -1)
                {
                    index = b.AddPropertyExtension2(val);
                    WindowLog.EcrireF("Ajoute {0}", val);
                }
                else
                {
                    WindowLog.EcrireF("Index {0}", index);
                    Object r = b.GetPropertyExtension2(index - 1);
                    WindowLog.EcrireF("Resultat {0}", r);
                }


            }
            catch (Exception e) { this.LogMethode(new Object[] { e }); }

        }

        //protected override void Command()
        //{
        //    try
        //    {
        //        ModelDoc2 mdl = App.ModelDoc2;
        //        List<Face2> ListeFaces = mdl.eSelect_RecupererListeObjets<Face2>();

        //        Body2 CorpsBase = ListeFaces[0].GetBody();
        //        Body2 CorpsTest = ListeFaces[1].GetBody();
        //        WindowLog.Ecrire(CorpsBase.Name);

        //        MathTransform mt = null;
        //        Boolean r = CorpsBase.GetCoincidenceTransform2((Object)CorpsTest, out mt);
        //        if (r == true)
        //            WindowLog.Ecrire("Coincidence : " + CorpsBase.Name + " et " + CorpsTest.Name);
        //        else
        //            WindowLog.Ecrire("Ne coincide pas : " + CorpsBase.Name + " et " + CorpsTest.Name);

        //    }
        //    catch (Exception e) { this.LogMethode(new Object[] { e }); }

        //}

        //protected override void Command()
        //{
        //    try
        //    {
        //        ModelDoc2 mdl = App.ModelDoc2;
        //        //var DossierExport = mdl.eDossier();
        //        //var NomFichier = mdl.eNomSansExt();

        //        Body2 Barre = null;

        //        var Face = mdl.eSelect_RecupererObjet<Face2>(1);

        //        if (Face.IsNull())
        //            Barre = mdl.eSelect_RecupererObjet<Body2>(1);
        //        else
        //            Barre = Face.GetBody();

        //        mdl.eEffacerSelection();

        //        WindowLog.Ecrire("Nom du corps : " + Barre.Name);

        //        var b = new AnalyseBarre(Barre, mdl);

        //        foreach (var u in b.ListeFaceUsinageExtremite)
        //        {
        //            WindowLog.Ecrire(u.LgUsinage * 1000);
        //            foreach (var e in u.ListeArreteDecoupe)
        //                e.eSelectEntite(true);
        //        }

        //        //foreach (var u in b.ListeFaceUsinageSection)
        //        //{
        //        //    WindowLog.Ecrire(u.LgUsinage * 1000);
        //        //    foreach (var e in u.ListeArreteDecoupe)
        //        //        e.eSelectEntite(true);
        //        //}
        //    }
        //    catch (Exception e) { this.LogMethode(new Object[] { e }); }

        //}
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
