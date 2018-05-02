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
        protected override void Command()
        {
            try
            {
                ModelDoc2 mdl = App.ModelDoc2;
                //var DossierExport = mdl.eDossier();
                //var NomFichier = mdl.eNomSansExt();

                Double NbTrou = 0;
                Double LgTotal = 0;

                foreach (var Face in mdl.eSelect_RecupererListeObjets<Face2>(-1))
                {
                    var Liste = Face.eListeDesBoucles();

                    NbTrou += Liste.Count - 1;

                    Double Lg = 0;
                    foreach (var Loop in Liste)
                    {
                        Loop.IsOuter();
                        Double LgLoop = 0;
                        foreach (var Coedge in Loop.eListeDesCoArrete())
                        {
                            Curve Courbe = ((Edge)Coedge.GetEdge()).GetCurve();
                            double Start, End; bool Ferme, Periodic;
                            Courbe.GetEndParams(out Start, out End, out Ferme, out Periodic);
                            LgLoop += Courbe.GetLength3(Start, End);
                        }

                        if (!Loop.IsOuter())
                            Lg += LgLoop;

                        LgLoop = 0;
                    }

                    LgTotal += Lg;
                }

                WindowLog.Ecrire(NbTrou);
                WindowLog.Ecrire(Math.Round(LgTotal * 1000));


                //var ListeNomConfigs = mdl.eListeNomConfiguration(eTypeConfig.Pliee);
                //ListeNomConfigs.Sort(new WindowsStringComparer());

                //for (int noCfg = 0; noCfg < ListeNomConfigs.Count; noCfg++)
                //{
                //    mdl.ClearSelection2(true);

                //    var NomConfigPliee = ListeNomConfigs[noCfg];
                //    mdl.ShowConfiguration2(NomConfigPliee);
                //    mdl.EditRebuild3();
                //    PartDoc Piece = mdl.ePartDoc();

                //    ListPID<Feature> ListeDossier = Piece.eListePIDdesFonctionsDePiecesSoudees(null);

                //    for (int noD = 0; noD < ListeDossier.Count; noD++)
                //    {
                //        Feature f = ListeDossier[noD];
                //        BodyFolder dossier = f.GetSpecificFeature2();

                //        if (dossier.eEstExclu() || dossier.IsNull() || (dossier.GetBodyCount() == 0)) continue;

                //        String Profil = dossier.eProp(CONSTANTES.PROFIL_NOM);
                //        String Longueur = dossier.eProp(CONSTANTES.PROFIL_LONGUEUR);

                //        if (String.IsNullOrWhiteSpace(Profil) || String.IsNullOrWhiteSpace(Longueur))
                //        {
                //            WindowLog.Ecrire("      Pas de barres");
                //            continue;
                //        }

                //        var Barre = dossier.ePremierCorps();

                //        foreach (var Face in Barre.eListeDesFaces())
                //        {

                //            foreach (var Loop in Face.eListeDesBoucles())
                //            {
                //                WindowLog.Ecrire("");
                //            }
                //        } 
                //    }
                //}

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
