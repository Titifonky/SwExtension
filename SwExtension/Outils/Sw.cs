using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;
using LogDebugging;
using System.Collections;

namespace Outils
{
    [Flags]
    public enum eTypeDoc
    {
        [Intitule("Pièce"), ExtFichier(".SLDPRT"), ExtGabarit(".PRTDOT")]
        Piece = 1,
        [Intitule("Assemblage"), ExtFichier(".SLDASM"), ExtGabarit(".ASMDOT")]
        Assemblage = 2,
        [Intitule("Dessin"), ExtFichier(".SLDDRW"), ExtGabarit(".DRWDOT")]
        Dessin = 4,
        [Intitule("Inconnu"), ExtFichier(".*")]
        Inconnu = 64
    }

    [Flags]
    public enum eTypeFichierExport
    {
        [Intitule("Pièce"), ExtFichier(".SLDPRT")]
        Piece = 1,
        [Intitule("Parasolid"), ExtFichier(".x_t")]
        Parasolid = 2,
        [Intitule("Parasolid binaire"), ExtFichier(".x_b")]
        ParasolidBinary = 4,
        [Intitule("Dxf"), ExtFichier(".dxf")]
        DXF = 8,
        [Intitule("Dwg"), ExtFichier(".dwg")]
        DWG = 16,
        [Intitule("Pdf"), ExtFichier(".pdf")]
        PDF = 32,
        [Intitule("Step"), ExtFichier(".stp")]
        STEP = 64
    }

    [Flags]
    public enum eTypeConfig
    {
        Racine = 1,
        DeBase = 2,
        Derivee = 4,
        Depliee = 8,
        Pliee = 16,
        Tous = Racine | DeBase | Derivee | Depliee | Pliee
    }

    [Flags]
    public enum eMatPropriete
    {
        [Intitule("Module d'élasticité"), TagXml("EX"), Unite("N/m²")]
        module_elastique = 1,
        [Intitule("Coefficient de Poisson"), TagXml("NUXY"), Unite("S.O.")]
        coef_poisson = 2,
        [Intitule("Module de cisaillement"), TagXml("GXY"), Unite("N/m²")]
        module_cisaillement = 4,
        [Intitule("Coefficient de dilatation thermique"), TagXml("ALPX"), Unite("mm/K")]
        coef_dilatation_thermique = 8,
        [Intitule("Masse volumique"), TagXml("DENS"), Unite("kg/m³")]
        densite = 16,
        [Intitule("Conductivité thermique"), TagXml("KX"), Unite("W/(m.K)")]
        conductivite_thermique = 32,
        [Intitule("Chaleur spécifique"), TagXml("C"), Unite("J/(kg.K)")]
        chaleur_specifique = 64,
        [Intitule("Limite de traction"), TagXml("SIGXT"), Unite("N/m²")]
        limite_traction = 128,
        [Intitule("Limite d'élasticité"), TagXml("SIGYLD"), Unite("N/m²")]
        limite_elastique = 256
    }

    public enum eOrientation
    {
        [Intitule("Portrait")]
        Portrait = 0,
        [Intitule("Paysage")]
        Paysage = 1
    }

    public enum eFormat
    {
        [Intitule("A0")]
        A0 = 5,
        [Intitule("A1")]
        A1 = 6,
        [Intitule("A2")]
        A2 = 7,
        [Intitule("A3")]
        A3 = 8,
        [Intitule("A4")]
        A4 = 9,
        [Intitule("A5")]
        A5 = 10,
        [Intitule("Utilisateur")]
        Utilisateur = 11
    }

    public enum eDxfFormat
    {
        [Intitule("R12")]
        R12 = 0,
        [Intitule("R13")]
        R13 = 1,
        [Intitule("R14")]
        R14 = 2,
        [Intitule("R2000")]
        R2000 = 3,
        [Intitule("R2004")]
        R2004 = 4,
        [Intitule("R2007")]
        R2007 = 5,
        [Intitule("R2010")]
        R2010 = 6,
        [Intitule("R2013")]
        R2013 = 7
    }

    [Flags]
    public enum eTypeCorps
    {
        Tole = 1,
        Barre = 2,
        Autre = 4,
        Tous = Tole | Barre | Autre
    }

    public enum eFeatureType
    {
        [Intitule("ExplodeLineProfileFeature")]
        swTnExplodeLineProfileFeature,

        [Intitule("InContextFeatHolder")]
        swTnInContextFeatHolder,

        [Intitule("MateCoincident")]
        swTnMateCoincident,

        [Intitule("MateConcentric")]
        swTnMateConcentric,

        [Intitule("MateDistanceDim")]
        swTnMateDistanceDim,

        [Intitule("MateGroup")]
        swTnMateGroup,

        [Intitule("MateInPlace")]
        swTnMateInPlace,

        [Intitule("MateParallel")]
        swTnMateParallel,

        [Intitule("MatePerpendicular")]
        swTnMatePerpendicular,

        [Intitule("MatePlanarAngleDim")]
        swTnMatePlanarAngleDim,

        [Intitule("MateSymmetric")]
        swTnMateSymmetric,

        [Intitule("MateTangent")]
        swTnMateTangent,

        [Intitule("MateWidth")]
        swTnMateWidth,

        [Intitule("Reference")]
        swTnReference,

        [Intitule("SmartComponentFeature")]
        swTnSmartComponentFeature,

        [Intitule("BaseBody")]
        swTnBaseBody,

        [Intitule("Blend")]
        swTnBlend,

        [Intitule("BlendCut")]
        swTnBlendCut,

        [Intitule("Boss")]
        swTnBoss,

        [Intitule("BossThin")]
        swTnBossThin,

        [Intitule("Cavity")]
        swTnCavity,

        [Intitule("Chamfer")]
        swTnChamfer,

        [Intitule("CirPattern")]
        swTnCirPattern,

        [Intitule("CombineBodies")]
        swTnCombineBodies,

        [Intitule("CosmeticThread")]
        swTnCosmeticThread,

        [Intitule("CurvePattern")]
        swTnCurvePattern,

        [Intitule("Cut")]
        swTnCut,

        [Intitule("CutThin")]
        swTnCutThin,

        [Intitule("Deform")]
        swTnDeform,

        [Intitule("DeleteBody")]
        swTnDeleteBody,

        [Intitule("DelFace")]
        swTnDelFace,

        [Intitule("DerivedCirPattern")]
        swTnDerivedCirPattern,

        [Intitule("DerivedLPattern")]
        swTnDerivedLPattern,

        [Intitule("Dome")]
        swTnDome,

        [Intitule("Draft")]
        swTnDraft,

        [Intitule("Emboss")]
        swTnEmboss,

        [Intitule("Extrusion")]
        swTnExtrusion,

        [Intitule("Fillet")]
        swTnFillet,

        [Intitule("Round fillet corner")]
        swTnFilletCorner,

        [Intitule("Helix")]
        swTnHelix,

        [Intitule("HoleWzd")]
        swTnHoleWzd,

        [Intitule("Imported")]
        swTnImported,

        [Intitule("ICE")]
        swTnICE,

        [Intitule("LocalCirPattern")]
        swTnLocalCirPattern,

        [Intitule("LocalLPattern")]
        swTnLocalLPattern,

        [Intitule("LPattern")]
        swTnLPattern,

        [Intitule("MirrorPattern")]
        swTnMirrorPattern,

        [Intitule("MirrorSolid")]
        swTnMirrorSolid,

        [Intitule("MoveCopyBody")]
        swTnMoveCopyBody,

        [Intitule("ReplaceFace")]
        swTnReplaceFace,

        [Intitule("RevCut")]
        swTnRevCut,

        [Intitule("Revolution")]
        swTnRevolution,

        [Intitule("RevolutionThin")]
        swTnRevolutionThin,

        [Intitule("Shape")]
        swTnShape,

        [Intitule("Shell")]
        swTnShell,

        [Intitule("Split")]
        swTnSplit,

        [Intitule("Stock")]
        swTnStock,

        [Intitule("Sweep")]
        swTnSweep,

        [Intitule("SweepCut")]
        swTnSweepCut,

        [Intitule("TablePattern")]
        swTnTablePattern,

        [Intitule("Thicken")]
        swTnThicken,

        [Intitule("ThickenCut")]
        swTnThickenCut,

        [Intitule("VarFillet")]
        swTnVarFillet,

        [Intitule("VolSweep")]
        swTnVolSweep,

        [Intitule("VolSweepCut")]
        swTnVolSweepCut,

        [Intitule("MirrorStock")]
        swTnMirrorStock,

        [Intitule("AbsoluteView")]
        swTnAbsoluteView,

        [Intitule("AlignGroup")]
        swTnAlignGroup,

        [Intitule("AuxiliaryView")]
        swTnAuxiliaryView,

        [Intitule("BomTemplate")]
        swTnBomTableAnchor,

        [Intitule("BomFeat")]
        swTnBomTableFeature,

        [Intitule("BomTemplate")]
        swTnBomTemplate,

        [Intitule("BreakLine")]
        swTnBreakLine,

        [Intitule("CenterMark")]
        swTnCenterMark,

        [Intitule("DetailCircle")]
        swTnDetailCircle,

        [Intitule("DetailView")]
        swTnDetailView,

        [Intitule("DrBreakoutSectionLine")]
        swTnDrBreakoutSectionLine,

        [Intitule("DrSectionLine")]
        swTnDrSectionLine,

        [Intitule("DrSheet")]
        swTnDrSheet,

        [Intitule("DrTemplate")]
        swTnDrTemplate,

        [Intitule("DrViewDetached")]
        swTnDrViewDetached,

        [Intitule("GeneralTableAnchor")]
        swTnGeneralTableAnchor,

        [Intitule("HoleTableAnchor")]
        swTnHoleTableAnchor,

        [Intitule("HoleTableFeat")]
        swTnHoleTableFeat,

        [Intitule("HoleTableFeat")]
        swTnHoleTableFeature,

        [Intitule("LiveSection")]
        swTnLiveSection,

        [Intitule("RelativeView")]
        swTnRelativeView,

        [Intitule("RevisionTableAnchor")]
        swTnRevisionTableAnchor,

        [Intitule("RevisionTableFeat")]
        swTnRevisionTableFeature,

        [Intitule("Section")]
        swTnSection,

        [Intitule("SectionAssemView")]
        swTnSectionAssemView,

        [Intitule("SectionPartView")]
        swTnSectionPartView,

        [Intitule("SectionView")]
        swTnSectionView,

        [Intitule("UnfoldedView")]
        swTnUnfoldedView,

        [Intitule("BlockFolder")]
        swTnBlocksFolder,

        [Intitule("BrokenDerivedPartFolder")]
        swTnBrokenDerivedPartFolder,

        [Intitule("CommentsFolder")]
        swTnCommentsFolder,

        [Intitule("CutListFolder")]
        swTnCutListFolder,

        [Intitule("DocsFolder")]
        swTnDocsFolder,

        [Intitule("FeatSolidBodyFolder")]
        swTnFeatSolidBodyFolder,

        [Intitule("FeatSurfaceBodyFolder")]
        swTnFeatSurfaceBodyFolder,

        [Intitule("FtrFolder")]
        swTnFeatureFolder,

        [Intitule("FtrFolder")]
        swTnFtrFolder,

        [Intitule("GridDetailFolder")]
        swTnGridDetailFolder,

        [Intitule("InsertedFeatureFolder")]
        swTnInsertedFeatureFolder,

        [Intitule("LiveSectionFolder")]
        swTnLiveSectionFolder,

        [Intitule("MateReferenceGroupFolder")]
        swTnMateReferenceGroupFolder,

        [Intitule("MaterialFolder")]
        swTnMaterialFolder,

        [Intitule("PosGroupFolder")]
        swTnPosGroupFolder,

        [Intitule("ProfileFtrFolder")]
        swTnProfileFtrFolder,

        [Intitule("RefAxisFtrFolder")]
        swTnRefAxisFtrFolder,

        [Intitule("RefPlaneFtrFolder")]
        swTnRefPlaneFtrFolder,

        [Intitule("SolidBodyFolder")]
        swTnSolidBodyFolder,

        [Intitule("SmartComponentFolder")]
        swTnSmartComponentFolder,

        [Intitule("SmartComponentRefFolder")]
        swTnSmartComponentRefFolder,

        [Intitule("SubAtomFolder")]
        swTnSubAtomFolder,

        [Intitule("SubWeldFolder")]
        swTnSubWeldFolder,

        [Intitule("SurfaceBodyFolder")]
        swTnSurfaceBodyFolder,

        [Intitule("TableFolder")]
        swTnTableFolder,

        [Intitule("Attribute")]
        swTnAttribute,

        [Intitule("BlockDef")]
        swTnBlockDef,

        [Intitule("Comments")]
        swTnComments,

        [Intitule("ConfigBuilderFeature")]
        swTnConfigBuilderFeature,

        [Intitule("Configuration")]
        swTnConfiguration,

        [Intitule("CurveInFile")]
        swTnCurveInFile,

        [Intitule("DesignTableFeature")]
        swTnDesignTableFeature,

        [Intitule("DetailCabinet")]
        swTnDetailCabinet,

        [Intitule("EmbedLinkDoc")]
        swTnEmbedLinkDoc,

        [Intitule("GridFeature")]
        swTnGridFeature,

        [Intitule("Journal")]
        swTnJournal,

        [Intitule("LibraryFeature")]
        swTnLibraryFeature,

        [Intitule("PartConfiguration")]
        swTnPartConfiguration,

        [Intitule("ReferenceBrowser")]
        swTnReferenceBrowser,

        [Intitule("ReferenceEmbedded")]
        swTnReferenceEmbedded,

        [Intitule("ReferenceInternal")]
        swTnReferenceInternal,

        [Intitule("Scale")]
        swTnScale,

        [Intitule("ViewerBodyFeature")]
        swTnViewerBodyFeature,

        [Intitule("XMLRulesFeature")]
        swTnXMLRulesFeature,

        [Intitule("MoldCoreCavitySolids")]
        swTnMoldCoreCavitySolids,

        [Intitule("MoldPartingGeom")]
        swTnMoldPartingGeom,

        [Intitule("MoldPartingLine")]
        swTnMoldPartingLine,

        [Intitule("MoldShutOffSrf")]
        swTnMoldShutOffSrf,

        [Intitule("AEM3DContact")]
        swTnAEM3DContact,

        [Intitule("AEMGravity")]
        swTnAEMGravity,

        [Intitule("AEMLinearDamper")]
        swTnAEMLinearDamper,

        [Intitule("AEMLinearForce")]
        swTnAEMLinearForce,

        [Intitule("AEMLinearMotor")]
        swTnAEMLinearMotor,

        [Intitule("AEMLinearSpring")]
        swTnAEMLinearSpring,

        [Intitule("AEMRotationalMotor")]
        swTnAEMRotationalMotor,

        [Intitule("AEMSimFeature")]
        swTnAEMSimFeature,

        [Intitule("AEMTorque")]
        swTnAEMTorque,

        [Intitule("AEMTorsionalDamper")]
        swTnAEMTorsionalDamper,

        [Intitule("AEMTorsionalSpring")]
        swTnAEMTorsionalSpring,

        [Intitule("CoordSys")]
        swTnCoordinateSystem,

        [Intitule("CoordSys")]
        swTnCoordSys,

        [Intitule("RefAxis")]
        swTnRefAxis,

        [Intitule("ReferenceCurve")]
        swTnReferenceCurve,

        [Intitule("RefPlane")]
        swTnRefPlane,

        [Intitule("AmbientLight")]
        swTnAmbientLight,

        [Intitule("CameraFeature")]
        swTnCameraFeature,

        [Intitule("DirectionLight")]
        swTnDirectionLight,

        [Intitule("PointLight")]
        swTnPointLight,

        [Intitule("SpotLight")]
        swTnSpotLight,

        [Intitule("SMBaseFlange")]
        swTnBaseFlange,

        [Intitule("Bending")]
        swTnBending,

        [Intitule("BreakCorner")]
        swTnBreakCorner,

        [Intitule("CornerTrim")]
        swTnCornerTrim,

        [Intitule("CrossBreak")]
        swTnCrossBreak,

        [Intitule("EdgeFlange")]
        swTnEdgeFlange,

        [Intitule("FlatPattern")]
        swTnFlatPattern,

        [Intitule("FlattenBends")]
        swTnFlattenBends,

        [Intitule("Fold")]
        swTnFold,

        [Intitule("FormToolInstance")]
        swTnFormToolInstance,

        [Intitule("Hem")]
        swTnHem,

        [Intitule("LoftedBend")]
        swTnLoftedBend,

        [Intitule("OneBend")]
        swTnOneBend,

        [Intitule("ProcessBends")]
        swTnProcessBends,

        [Intitule("SheetMetal")]
        swTnSheetMetal,

        [Intitule("SketchBend")]
        swTnSketchBend,

        [Intitule("SM3dBend")]
        swTnSM3dBend,

        [Intitule("SMBaseFlange")]
        swTnSMBaseFlange,

        [Intitule("SMMiteredFlange")]
        swTnSMMiteredFlange,

        [Intitule("ToroidalBend")]
        swTnToroidalBend,

        [Intitule("UiBend")]
        swTnUiBend,

        [Intitule("UnFold")]
        swTnUnFold,

        [Intitule("3DProfileFeature")]
        swTn3DProfileFeature,

        [Intitule("3DSplineCurve")]
        swTn3DSplineCurve,

        [Intitule("CompositeCurve")]
        swTnCompositeCurve,

        [Intitule("LayoutProfileFeature")]
        swTnLayoutProfileFeature,

        [Intitule("PLine")]
        swTnPLine,

        [Intitule("ProfileFeature")]
        swTnProfileFeature,

        [Intitule("RefCurve")]
        swTnRefCurve,

        [Intitule("RefPoint")]
        swTnRefPoint,

        [Intitule("SketchBlockDef")]
        swTnSketchBlockDefinition,

        [Intitule("SketchHole")]
        swTnSketchHole,

        [Intitule("SketchPattern")]
        swTnSketchPattern,

        [Intitule("SketchBitmap")]
        swTnSketchPicture,

        [Intitule("BlendRefSurface")]
        swTnBlendRefSurface,

        [Intitule("FillRefSurface")]
        swTnFillRefSurface,

        [Intitule("MidRefSurface")]
        swTnMidRefSurface,

        [Intitule("OffsetRefSurface")]
        swTnOffsetRefSurface,

        [Intitule("RadiateRefSurface")]
        swTnRadiateRefSurface,

        [Intitule("RefSurface")]
        swTnRefSurface,

        [Intitule("RevolvRefSurf")]
        swTnRevolvRefSurf,

        [Intitule("RuledSrfFromEdge")]
        swTnRuledSrfFromEdge,

        [Intitule("SewRefSurface")]
        swTnSewRefSurface,

        [Intitule("SurfCut")]
        swTnSurfCut,

        [Intitule("SweepRefSurface")]
        swTnSweepRefSurface,

        [Intitule("TrimRefSurface")]
        swTnTrimRefSurface,

        [Intitule("UnTrimRefSurf")]
        swTnUnTrimRefSurf,

        [Intitule("VolSweepRefSurface")]
        swTnVolSweepRefSurface,

        [Intitule("Gusset")]
        swTnGusset,

        [Intitule("WeldMemberFeat")]
        swTnWeldMemberFeat,

        [Intitule("WeldmentFeature")]
        swTnWeldmentFeature,

        [Intitule("WeldmentTableAnchor")]
        swTnWeldmentTableAnchor,

        [Intitule("WeldmentTableFeat")]
        swTnWeldmentTableFeature
    }

    public static class FeatureType
    {
        public const String swTnExplodeLineProfileFeature = "ExplodeLineProfileFeature";
        public const String swTnInContextFeatHolder = "InContextFeatHolder";
        public const String swTnMateCoincident = "MateCoincident";
        public const String swTnMateConcentric = "MateConcentric";
        public const String swTnMateDistanceDim = "MateDistanceDim";
        public const String swTnMateGroup = "MateGroup";
        public const String swTnMateInPlace = "MateInPlace";
        public const String swTnMateParallel = "MateParallel";
        public const String swTnMatePerpendicular = "MatePerpendicular";
        public const String swTnMatePlanarAngleDim = "MatePlanarAngleDim";
        public const String swTnMateSymmetric = "MateSymmetric";
        public const String swTnMateTangent = "MateTangent";
        public const String swTnMateWidth = "MateWidth";
        public const String swTnReference = "Reference";
        public const String swTnSmartComponentFeature = "SmartComponentFeature";
        public const String swTnBaseBody = "BaseBody";
        public const String swTnBlend = "Blend";
        public const String swTnBlendCut = "BlendCut";
        public const String swTnBoss = "Boss";
        public const String swTnBossThin = "BossThin";
        public const String swTnCavity = "Cavity";
        public const String swTnChamfer = "Chamfer";
        public const String swTnCirPattern = "CirPattern";
        public const String swTnCombineBodies = "CombineBodies";
        public const String swTnCosmeticThread = "CosmeticThread";
        public const String swTnCurvePattern = "CurvePattern";
        public const String swTnCut = "Cut";
        public const String swTnCutThin = "CutThin";
        public const String swTnDeform = "Deform";
        public const String swTnDeleteBody = "DeleteBody";
        public const String swTnDelFace = "DelFace";
        public const String swTnDerivedCirPattern = "DerivedCirPattern";
        public const String swTnDerivedLPattern = "DerivedLPattern";
        public const String swTnDome = "Dome";
        public const String swTnDraft = "Draft";
        public const String swTnEmboss = "Emboss";
        public const String swTnExtrusion = "Extrusion";
        public const String swTnFillet = "Fillet";
        public const String swTnFilletCorner = "Round fillet corner";
        public const String swTnHelix = "Helix";
        public const String swTnHoleWzd = "HoleWzd";
        public const String swTnImported = "Imported";
        public const String swTnICE = "ICE";
        public const String swTnLocalCirPattern = "LocalCirPattern";
        public const String swTnLocalLPattern = "LocalLPattern";
        public const String swTnLPattern = "LPattern";
        public const String swTnMirrorPattern = "MirrorPattern";
        public const String swTnMirrorSolid = "MirrorSolid";
        public const String swTnMoveCopyBody = "MoveCopyBody";
        public const String swTnReplaceFace = "ReplaceFace";
        public const String swTnRevCut = "RevCut";
        public const String swTnRevolution = "Revolution";
        public const String swTnRevolutionThin = "RevolutionThin";
        public const String swTnShape = "Shape";
        public const String swTnShell = "Shell";
        public const String swTnSplit = "Split";
        public const String swTnStock = "Stock";
        public const String swTnSweep = "Sweep";
        public const String swTnSweepCut = "SweepCut";
        public const String swTnTablePattern = "TablePattern";
        public const String swTnThicken = "Thicken";
        public const String swTnThickenCut = "ThickenCut";
        public const String swTnVarFillet = "VarFillet";
        public const String swTnVolSweep = "VolSweep";
        public const String swTnVolSweepCut = "VolSweepCut";
        public const String swTnMirrorStock = "MirrorStock";
        public const String swTnAbsoluteView = "AbsoluteView";
        public const String swTnAlignGroup = "AlignGroup";
        public const String swTnAuxiliaryView = "AuxiliaryView";
        public const String swTnBomTableAnchor = "BomTemplate";
        public const String swTnBomTableFeature = "BomFeat";
        public const String swTnBomTemplate = "BomTemplate";
        public const String swTnBreakLine = "BreakLine";
        public const String swTnCenterMark = "CenterMark";
        public const String swTnDetailCircle = "DetailCircle";
        public const String swTnDetailView = "DetailView";
        public const String swTnDrBreakoutSectionLine = "DrBreakoutSectionLine";
        public const String swTnDrSectionLine = "DrSectionLine";
        public const String swTnDrSheet = "DrSheet";
        public const String swTnDrTemplate = "DrTemplate";
        public const String swTnDrViewDetached = "DrViewDetached";
        public const String swTnGeneralTableAnchor = "GeneralTableAnchor";
        public const String swTnHoleTableAnchor = "HoleTableAnchor";
        public const String swTnHoleTableFeat = "HoleTableFeat";
        public const String swTnHoleTableFeature = "HoleTableFeat";
        public const String swTnLiveSection = "LiveSection";
        public const String swTnRelativeView = "RelativeView";
        public const String swTnRevisionTableAnchor = "RevisionTableAnchor";
        public const String swTnRevisionTableFeature = "RevisionTableFeat";
        public const String swTnSection = "Section";
        public const String swTnSectionAssemView = "SectionAssemView";
        public const String swTnSectionPartView = "SectionPartView";
        public const String swTnSectionView = "SectionView";
        public const String swTnUnfoldedView = "UnfoldedView";
        public const String swTnBlocksFolder = "BlockFolder";
        public const String swTnBrokenDerivedPartFolder = "BrokenDerivedPartFolder";
        public const String swTnCommentsFolder = "CommentsFolder";
        public const String swTnCutListFolder = "CutListFolder";
        public const String swTnDocsFolder = "DocsFolder";
        public const String swTnFeatSolidBodyFolder = "FeatSolidBodyFolder";
        public const String swTnFeatSurfaceBodyFolder = "FeatSurfaceBodyFolder";
        public const String swTnFeatureFolder = "FtrFolder";
        public const String swTnFtrFolder = "FtrFolder";
        public const String swTnGridDetailFolder = "GridDetailFolder";
        public const String swTnInsertedFeatureFolder = "InsertedFeatureFolder";
        public const String swTnLiveSectionFolder = "LiveSectionFolder";
        public const String swTnMateReferenceGroupFolder = "MateReferenceGroupFolder";
        public const String swTnMaterialFolder = "MaterialFolder";
        public const String swTnPosGroupFolder = "PosGroupFolder";
        public const String swTnProfileFtrFolder = "ProfileFtrFolder";
        public const String swTnRefAxisFtrFolder = "RefAxisFtrFolder";
        public const String swTnRefPlaneFtrFolder = "RefPlaneFtrFolder";
        public const String swTnSolidBodyFolder = "SolidBodyFolder";
        public const String swTnSmartComponentFolder = "SmartComponentFolder";
        public const String swTnSmartComponentRefFolder = "SmartComponentRefFolder";
        public const String swTnSubAtomFolder = "SubAtomFolder";
        public const String swTnSubWeldFolder = "SubWeldFolder";
        public const String swTnSurfaceBodyFolder = "SurfaceBodyFolder";
        public const String swTnTableFolder = "TableFolder";
        public const String swTnAttribute = "Attribute";
        public const String swTnBlockDef = "BlockDef";
        public const String swTnComments = "Comments";
        public const String swTnConfigBuilderFeature = "ConfigBuilderFeature";
        public const String swTnConfiguration = "Configuration";
        public const String swTnCurveInFile = "CurveInFile";
        public const String swTnDesignTableFeature = "DesignTableFeature";
        public const String swTnDetailCabinet = "DetailCabinet";
        public const String swTnEmbedLinkDoc = "EmbedLinkDoc";
        public const String swTnGridFeature = "GridFeature";
        public const String swTnJournal = "Journal";
        public const String swTnLibraryFeature = "LibraryFeature";
        public const String swTnPartConfiguration = "PartConfiguration";
        public const String swTnReferenceBrowser = "ReferenceBrowser";
        public const String swTnReferenceEmbedded = "ReferenceEmbedded";
        public const String swTnReferenceInternal = "ReferenceInternal";
        public const String swTnScale = "Scale";
        public const String swTnViewerBodyFeature = "ViewerBodyFeature";
        public const String swTnXMLRulesFeature = "XMLRulesFeature";
        public const String swTnMoldCoreCavitySolids = "MoldCoreCavitySolids";
        public const String swTnMoldPartingGeom = "MoldPartingGeom";
        public const String swTnMoldPartingLine = "MoldPartingLine";
        public const String swTnMoldShutOffSrf = "MoldShutOffSrf";
        public const String swTnAEM3DContact = "AEM3DContact";
        public const String swTnAEMGravity = "AEMGravity";
        public const String swTnAEMLinearDamper = "AEMLinearDamper";
        public const String swTnAEMLinearForce = "AEMLinearForce";
        public const String swTnAEMLinearMotor = "AEMLinearMotor";
        public const String swTnAEMLinearSpring = "AEMLinearSpring";
        public const String swTnAEMRotationalMotor = "AEMRotationalMotor";
        public const String swTnAEMSimFeature = "AEMSimFeature";
        public const String swTnAEMTorque = "AEMTorque";
        public const String swTnAEMTorsionalDamper = "AEMTorsionalDamper";
        public const String swTnAEMTorsionalSpring = "AEMTorsionalSpring";
        public const String swTnCoordinateSystem = "CoordSys";
        public const String swTnCoordSys = "CoordSys";
        public const String swTnRefAxis = "RefAxis";
        public const String swTnReferenceCurve = "ReferenceCurve";
        public const String swTnRefPlane = "RefPlane";
        public const String swTnAmbientLight = "AmbientLight";
        public const String swTnCameraFeature = "CameraFeature";
        public const String swTnDirectionLight = "DirectionLight";
        public const String swTnPointLight = "PointLight";
        public const String swTnSpotLight = "SpotLight";
        public const String swTnBaseFlange = "SMBaseFlange";
        public const String swTnBending = "Bending";
        public const String swTnBreakCorner = "BreakCorner";
        public const String swTnCornerTrim = "CornerTrim";
        public const String swTnCrossBreak = "CrossBreak";
        public const String swTnEdgeFlange = "EdgeFlange";
        public const String swTnFlatPattern = "FlatPattern";
        public const String swTnFlattenBends = "FlattenBends";
        public const String swTnFold = "Fold";
        public const String swTnFormToolInstance = "FormToolInstance";
        public const String swTnHem = "Hem";
        public const String swTnLoftedBend = "LoftedBend";
        public const String swTnOneBend = "OneBend";
        public const String swTnProcessBends = "ProcessBends";
        public const String swTnSheetMetal = "SheetMetal";
        public const String swTnSketchBend = "SketchBend";
        public const String swTnSM3dBend = "SM3dBend";
        public const String swTnSMBaseFlange = "SMBaseFlange";
        public const String swTnSMMiteredFlange = "SMMiteredFlange";
        public const String swTnToroidalBend = "ToroidalBend";
        public const String swTnUiBend = "UiBend";
        public const String swTnUnFold = "UnFold";
        public const String swTn3DProfileFeature = "3DProfileFeature";
        public const String swTn3DSplineCurve = "3DSplineCurve";
        public const String swTnCompositeCurve = "CompositeCurve";
        public const String swTnLayoutProfileFeature = "LayoutProfileFeature";
        public const String swTnPLine = "PLine";
        public const String swTnProfileFeature = "ProfileFeature";
        public const String swTnRefCurve = "RefCurve";
        public const String swTnRefPoint = "RefPoint";
        public const String swTnSketchBlockDefinition = "SketchBlockDef";
        public const String swTnSketchHole = "SketchHole";
        public const String swTnSketchPattern = "SketchPattern";
        public const String swTnSketchPicture = "SketchBitmap";
        public const String swTnBlendRefSurface = "BlendRefSurface";
        public const String swTnFillRefSurface = "FillRefSurface";
        public const String swTnMidRefSurface = "MidRefSurface";
        public const String swTnOffsetRefSurface = "OffsetRefSurface";
        public const String swTnRadiateRefSurface = "RadiateRefSurface";
        public const String swTnRefSurface = "RefSurface";
        public const String swTnRevolvRefSurf = "RevolvRefSurf";
        public const String swTnRuledSrfFromEdge = "RuledSrfFromEdge";
        public const String swTnSewRefSurface = "SewRefSurface";
        public const String swTnSurfCut = "SurfCut";
        public const String swTnSweepRefSurface = "SweepRefSurface";
        public const String swTnTrimRefSurface = "TrimRefSurface";
        public const String swTnUnTrimRefSurf = "UnTrimRefSurf";
        public const String swTnVolSweepRefSurface = "VolSweepRefSurface";
        public const String swTnGusset = "Gusset";
        public const String swTnWeldMemberFeat = "WeldMemberFeat";
        public const String swTnWeldmentFeature = "WeldmentFeature";
        public const String swTnWeldmentTableAnchor = "WeldmentTableAnchor";
        public const String swTnWeldmentTableFeature = "WeldmentTableFeat";
    }

    public static class swSelectType
    {
        public const String swSelLOCATIONS = "LOCATIONS";
        public const String swSelUNSUPPORTED = "UNSUPPORTED";
        public const String swSelDISPLAYSTATE = "VISUALSTATE";
        public const String swSelANNOTATIONTABLES = "ANNOTATIONTABLES";
        public const String swSelANNOTATIONVIEW = "ANNVIEW";
        public const String swSelATTRIBUTES = "ATTRIBUTE";
        public const String swSelDATUMAXES = "AXIS";
        public const String swSelBODYFOLDER = "BDYFOLDER";
        public const String swSelBLOCKDEF = "BLOCKDEF";
        public const String swSelBODYFEATURES = "BODYFEATURE";
        public const String swSelBOMS = "BOM";
        public const String swSelBOMFEATURES = "BOMFEATURE";
        public const String swSelBOMTEMPS = "BOMTEMP";
        public const String swSelBREAKLINES = "BREAKLINE";
        public const String swSelBROWSERITEM = "BROWSERITEM";
        public const String swSelCAMERAS = "CAMERAS";
        public const String swSelCENTERLINES = "CENTERLINE";
        public const String swSelCENTERMARKS = "CENTERMARKS";
        public const String swSelCENTERMARKSYMS = "CENTERMARKSYMS";
        public const String swSelCOMMENT = "COMMENT";
        public const String swSelCOMMENTSFOLDER = "COMMENTSFOLDER";
        public const String swSelCOMPONENTS = "COMPONENT";
        public const String swSelCOMPPATTERN = "COMPPATTERN";
        public const String swSelCONFIGURATIONS = "CONFIGURATIONS";
        public const String swSelCONNECTIONPOINTS = "CONNECTIONPOINT";
        public const String swSelCOORDSYS = "COORDSYS";
        public const String swSelCOSMETICWELDS = "COSMETICWELDS";
        public const String swSelCTHREADS = "CTHREAD";
        public const String swSelDATUMPOINTS = "DATUMPOINT";
        public const String swSelDATUMTAGS = "DATUMTAG";
        public const String swSelDCABINETS = "DCABINET";
        public const String swSelDETAILCIRCLES = "DETAILCIRCLE";
        public const String swSelDIMENSIONS = "DIMENSION";
        public const String swSelDOCSFOLDER = "DOCSFOLDER";
        public const String swSelDOWELSYMS = "DOWLELSYM";
        public const String swSelDRAWINGVIEWS = "DRAWINGVIEW";
        public const String swSelDTMTARGS = "DTMTARG";
        public const String swSelEDGES = "EDGE";
        public const String swSelEMBEDLINKDOC = "EMBEDLINKDOC";
        public const String swSelEQNFOLDER = "EQNFOLDER";
        public const String swSelEXPLLINES = "EXPLODELINES";
        public const String swSelEXPLSTEPS = "EXPLODESTEPS";
        public const String swSelEXPLVIEWS = "EXPLODEVIEWS";
        public const String swSelEXTSKETCHPOINTS = "EXTSKETCHPOINT";
        public const String swSelEXTSKETCHSEGS = "EXTSKETCHSEGMENT";
        public const String swSelEXTSKETCHTEXT = "EXTSKETCHTEXT";
        public const String swSelFACES = "FACE";
        public const String swSelFRAMEPOINT = "FRAMEPOINT";
        public const String swSelFTRFOLDER = "FTRFOLDER";
        public const String swSelGENERALTABLEFEAT = "GENERALTABLEFEAT";
        public const String swSelGTOLS = "GTOL";
        public const String swSelHELIX = "HELIX";
        public const String swSelHOLETABLEFEATS = "HOLETABLE";
        public const String swSelHOLETABLEAXES = "HOLETABLEAXIS";
        public const String swSelVIEWERHYPERLINK = "HYPERLINK";
        public const String swSelIMPORTFOLDER = "IMPORTFOLDER";
        public const String swSelINCONTEXTFEAT = "INCONTEXTFEAT";
        public const String swSelINCONTEXTFEATS = "INCONTEXTFEATS";
        public const String swSelJOURNAL = "JOURNAL";
        public const String swSelLEADERS = "LEADER";
        public const String swSelLIGHTS = "LIGHTS";
        public const String swSelMAGNETICLINES = "MAGNETICLINES";
        public const String swSelMANIPULATORS = "MANIPULATOR";
        public const String swSelMATES = "MATE";
        public const String swSelMATEGROUP = "MATEGROUP";
        public const String swSelMATEGROUPS = "MATEGROUPS";
        public const String swSelMATESUPPLEMENT = "MATESUPPLEMENT";
        public const String swSelNOTES = "NOTE";
        public const String swSelOBJGROUP = "OBJGROUP";
        public const String swSelOLEITEMS = "OLEITEM";
        public const String swSelPICTUREBODIES = "PICTURE BODY";
        public const String swSelDATUMPLANES = "PLANE";
        public const String swSelPOINTREFS = "POINTREF";
        public const String swSelPOSGROUP = "POSGROUP";
        public const String swSelPUNCHTABLEFEATS = "PUNCHTABLE";
        public const String swSelREFCURVES = "REFCURVE";
        public const String swSelREFERENCECURVES = "REFERENCECURVES";
        public const String swSelREFEDGES = "REFERENCE-EDGE";
        public const String swSelDATUMLINES = "REFLINE";
        public const String swSelREFSURFACES = "REFSURFACE";
        public const String swSelREVISIONTABLE = "REVISIONTABLE";
        public const String swSelREVISIONTABLEFEAT = "REVISIONTABLEFEAT";
        public const String swSelFABRICATEDROUTE = "ROUTEFABRICATED";
        public const String swSelROUTEPOINTS = "ROUTEPOINT";
        public const String swSelSECTIONLINES = "SECTIONLINE";
        public const String swSelSECTIONTEXT = "SECTIONTEXT";
        public const String swSelSELECTIONSETFOLDER = "SELECTIONSETFOLDER";
        public const String swSelSFSYMBOLS = "SFSYMBOL";
        public const String swSelSHEETS = "SHEET";
        public const String swSelSILHOUETTES = "SILHOUETTE";
        public const String swSelSIMULATION = "SIMULATION";
        public const String swSelSIMELEMENT = "SIMULATION_ELEMENT";
        public const String swSelSKETCHES = "SKETCH";
        public const String swSelSKETCHBITMAP = "SKETCHBITMAP";
        public const String swSelSKETCHCONTOUR = "SKETCHCONTOUR";
        public const String swSelSKETCHHATCH = "SKETCHHATCH";
        public const String swSelSKETCHPOINTS = "SKETCHPOINT";
        public const String swSelSKETCHPOINTFEAT = "SKETCHPOINTFEAT";
        public const String swSelSKETCHREGION = "SKETCHREGION";
        public const String swSelSKETCHSEGS = "SKETCHSEGMENT";
        public const String swSelSKETCHTEXT = "SKETCHTEXT";
        public const String swSelSOLIDBODIES = "SOLIDBODY";
        public const String swSelSELECTIONSETNODE = "SUBSELECTIONSETNODE";
        public const String swSelSUBSKETCHDEF = "SUBSKETCHDEF";
        public const String swSelSUBSKETCHINST = "SUBSKETCHINST";
        public const String swSelSUBWELDFOLDER = "SUBWELDMENT";
        public const String swSelSURFACEBODIES = "SURFACEBODY";
        public const String swSelSWIFTANNOTATIONS = "SWIFTANN";
        public const String swSelSWIFTFEATURES = "SWIFTFEATURE";
        public const String swSelSWIFTSCHEMA = "SWIFTSCHEMA";
        public const String swSelTITLEBLOCK = "TITLEBLOCK";
        public const String swSelTITLEBLOCKTABLEFEAT = "TITLEBLOCKTABLEFEAT";
        public const String swSelVERTICES = "VERTEX";
        public const String swSelARROWS = "VIEWARROW";
        public const String swSelWELDS = "WELD";
        public const String swSelWELDBEADS3 = "WELDBEADS";
        public const String swSelWELDMENT = "WELDMENT";
        public const String swSelWELDMENTTABLEFEATS = "WELDMENTTABLE";
        public const String swSelZONES = "ZONES";
    }

    internal static class CONSTANTES
    {
        internal const String CONFIG_DEPLIEE_PATTERN = "^([0-9]+)(SM-FLAT-PATTERN)((.)+)$";
        internal const String CONFIG_DEPLIEE = "SM-FLAT-PATTERN";
        internal const String CONFIG_PLIEE_PATTERN = @"^\d+$";
        internal const String ARTICLE_LISTE_DES_PIECES_SOUDEES = "Article-liste-des-pièces-soudées";
        internal const String EPAISSEUR_DE_TOLERIE = "Epaisseur de tôlerie";
        internal const String NO_DOSSIER = "NoDossier";
        internal const String REF_DOSSIER = "RefDossier";
        internal const String PREFIXE_REF_DOSSIER = "P";
        internal const String DESC_DOSSIER = "Description";
        internal const String NOM_DOSSIER = "NomDossier";
        internal const String NOM_ESQUISSE_NUMEROTER = "REPERAGE_DOSSIER";
        internal const String NOM_BLOCK_ESQUISSE_NUMEROTER = "REPERAGE_DOSSIER_BLOC.SLDBLK";
        internal const String NO_CONFIG = "NoConfigPliee";
        internal const String NOM_ELEMENT = "Element";
        internal const String PROFIL_NOM = "Profil";
        internal const String PROFIL_ANGLE1 = "ANGLE1";
        internal const String PROFIL_ANGLE2 = "ANGLE2";
        internal const String PROFIL_LONGUEUR = "LONGUEUR";
        internal const String PROFIL_MASSE = "Masse";
        internal const String PROFIL_MATERIAU = "MATERIAL";
        internal const String LIGNES_DE_PLIAGE = "Lignes de pliage";
        internal const String CUBE_DE_VISUALISATION = "Cube de visualisation";
        internal const String MODELE_DE_DESSIN_LASER = "MacroLaser";
        internal const String MODELE_DE_TABLE_DE_PLIAGE_LASER = "MacroLaser";
        internal const String NOM_CORPS_DEPLIEE = "Etat déplié";
        internal const String ETAT_D_AFFICHAGE = "Etat d'affichage-";
        internal const String MATERIAUX_NON_SPECIFIE = "Materiau non spécifié";
        internal const String DEPLIAGE = "Dépliage";
        internal const String DOSSIER_PROP_MASSE = "Masse";
        internal const String PROPRIETE_QUANTITE = "Quantite";
        internal const String PROPRIETE_NOCLIENT = "NoClient";
        internal const String PROPRIETE_NOCOMMANDE = "NoCommande";
        internal const String DOSSIER_DVP = "Plans Développées";
        internal const String DOSSIER_BARRE = "Export barres";
    }

    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]
    public class SwObjectPID<T>
        where T : class
    {
        private T _swObjet = null;
        private Object _PID = null;
        private ModelDoc2 _Mdl;

        public SwObjectPID(T swObjet, ModelDoc2 mdl)
        {
            _swObjet = swObjet;
            _Mdl = mdl;
            _PID = _Mdl.Extension.GetPersistReference3(swObjet);
        }

        public T Get
        {
            get
            {
                int Err = 0;

                T obj = (T)_Mdl.Extension.GetObjectByPersistReference3(_PID, out Err);

                Log.Message(Err);

                if ((int)swPersistReferencedObjectStates_e.swPersistReferencedObject_Ok == Err)
                    return obj;

                return null;
            }
        }

        public ModelDoc2 Modele
        {
            get { return _Mdl; }
        }

        public void Maj(ref T swObjet)
        {
            int Err = 0;

            Object obj = _Mdl.Extension.GetObjectByPersistReference3(_PID, out Err);

            if ((int)swPersistReferencedObjectStates_e.swPersistReferencedObject_Ok == Err)
                swObjet = obj as T;
        }

        public void Maj()
        {
            int Err = 0;

            Object obj = _Mdl.Extension.GetObjectByPersistReference3(_PID, out Err);

            if ((int)swPersistReferencedObjectStates_e.swPersistReferencedObject_Ok == Err)
                _swObjet = obj as T;
        }
    }

    /// <summary>
    /// Liste des objets Solidworks, seul le PID est stocké.
    /// La liste renvoi l'objet via ModelDoc2.GetObjectByPersistReference3()
    /// </summary>
    /// <typeparam name="T"></typeparam>
    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ListPID<T> : IEnumerable
        where T : class
    {
        private List<Object> _Liste;
        private ModelDoc2 _Mdl;

        public ListPID(ModelDoc2 mdl)
        {
            _Mdl = mdl;
            _Liste = new List<object>();
        }

        public ListPID(ModelDoc2 mdl, List<T> liste)
        {
            _Mdl = mdl;
            _Liste = new List<object>();

            foreach (T o in liste)
                Add(o);
        }

        public void Add(T o)
        {
            _Liste.Add(_Mdl.Extension.GetPersistReference3(o));
        }

        public void Remove(T o)
        {
            _Liste.Remove(_Mdl.Extension.GetPersistReference3(o));
        }

        public T this[int index]
        {
            get
            {
                int Err = 0;

                T pO = (T)_Mdl.Extension.GetObjectByPersistReference3(_Liste[index], out Err);

                Log.Message(Err);

                if ((int)swPersistReferencedObjectStates_e.swPersistReferencedObject_Ok == Err)
                    return pO;

                return null;
            }
        }

        public int Count
        {
            get
            {
                return _Liste.Count;
            }
        }

        public IEnumerator GetEnumerator()
        {
            return new Enumerator(this);
        }

        private class Enumerator : IEnumerator
        {
            private int _index = -1;
            private ListPID<T> _ListPID;

            public Enumerator(ListPID<T> t)
            {
                _ListPID = t;
            }

            public bool MoveNext()
            {
                if (_index < _ListPID._Liste.Count - 1)
                {
                    _index++;
                    return true;
                }
                else
                {
                    return false;
                }
            }

            public void Reset()
            {
                _index = -1;
            }

            public object Current
            {
                get
                {
                    return _ListPID[_index];
                }
            }
        }
    }

    public class EqualCompareComponent2 : IEqualityComparer<Component2>
    {
        bool IEqualityComparer<Component2>.Equals(Component2 x, Component2 y)
        {
            if (Sw.eIsSame(x, y))
                return true;

            return false;
        }

        int IEqualityComparer<Component2>.GetHashCode(Component2 obj)
        {
            ModelDoc2 Mdl = obj.GetModelDoc2();
            ModelDocExtension Ext = Mdl.Extension;

            return Ext.GetPersistReference3(obj);
        }
    }

    public class EqualCompareModelDoc2 : IEqualityComparer<ModelDoc2>
    {
        bool IEqualityComparer<ModelDoc2>.Equals(ModelDoc2 x, ModelDoc2 y)
        {
            if (x.GetPathName().ToUpperInvariant() == y.GetPathName().ToUpperInvariant())
                return true;

            return false;
        }

        int IEqualityComparer<ModelDoc2>.GetHashCode(ModelDoc2 obj)
        {
            return obj.GetPathName().GetHashCode();
        }
    }

    public class CompareModelDoc2 : IComparer<ModelDoc2>
    {
        private WindowsStringComparer sc = new WindowsStringComparer();

        int IComparer<ModelDoc2>.Compare(ModelDoc2 x, ModelDoc2 y)
        {
            return sc.Compare(x.eNomAvecExt(), y.eNomAvecExt());
        }
    }


    public static class App
    {

        static App()
        {
            MacroEnCours = false;
        }

        public static SldWorks Sw = null;
        public static ModelDoc2 ModelDoc2 { get { return Sw.ActiveDoc; } }
        public static AssemblyDoc AssemblyDoc { get { return Sw.ActiveDoc; } }
        public static PartDoc PartDoc { get { return Sw.ActiveDoc; } }
        public static DrawingDoc DrawingDoc { get { return Sw.ActiveDoc; } }

        public static Boolean MacroEnCours { get; set; }
    }

    public static class Sw
    {
        /// <summary>
        /// Retourne le chemin complet du bloc
        /// </summary>
        /// <param name="nomBloc">avec extension</param>
        /// <returns></returns>
        public static String CheminBloc(String nomBloc)
        {
            String cheminbloc = "";

            var CheminDossierBloc = App.Sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsBlocks).Split(new String[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var chemin in CheminDossierBloc)
            {
                var d = new DirectoryInfo(chemin);
                var r = d.GetFiles(nomBloc, SearchOption.AllDirectories);
                if (r.Length > 0)
                {
                    cheminbloc = r[0].FullName;
                    break;
                }
            }

            return cheminbloc;
        }

        //========================================================================================

        public static String eKeySansConfig(this Component2 cp)
        {
            return cp.GetPathName();
        }

        public static String eKeyAvecConfig(this Component2 cp)
        {
            return cp.eKeyAvecConfig(cp.eNomConfiguration());
        }

        public static String eKeyAvecConfig(this Component2 cp, String nomConfig)
        {
            return cp.GetPathName() + "__" + nomConfig;
        }

        //========================================================================================

        public static String eDescription(this eFeatureType type)
        {
            return type.GetEnumInfo<Intitule>();
        }

        public static swSelectType_e eGetSwSelectType_e<T>()
        {
            String N = typeof(T).Name.ToLower();

            if (N == "breakline") return swSelectType_e.swSelBREAKLINES;
            if (N == "comment") return swSelectType_e.swSelCOMMENT;
            if (N == "component2") return swSelectType_e.swSelCOMPONENTS;
            if (N == "datumtag") return swSelectType_e.swSelDATUMTAGS;
            if (N == "datumtargetsym") return swSelectType_e.swSelDTMTARGS;
            if (N == "dimxpertmanager") return swSelectType_e.swSelSWIFTSCHEMA;
            if (N == "displaydimension") return swSelectType_e.swSelDIMENSIONS;
            if (N == "dowelsymbol") return swSelectType_e.swSelDOWELSYMS;
            if (N == "edge") return swSelectType_e.swSelEDGES;
            if (N == "face2") return swSelectType_e.swSelFACES;
            if (N == "gtol") return swSelectType_e.swSelGTOLS;
            if (N == "mateloadreference") return swSelectType_e.swSelMATESUPPLEMENT;
            if (N == "note") return swSelectType_e.swSelNOTES;
            if (N == "projectionarrow") return swSelectType_e.swSelARROWS;
            if (N == "sfsymbol") return swSelectType_e.swSelSFSYMBOLS;
            if (N == "sheet") return swSelectType_e.swSelSHEETS;
            if (N == "silhouetteedge ") return swSelectType_e.swSelSILHOUETTES;
            if (N == "sketchblockinstance") return swSelectType_e.swSelSUBSKETCHINST;
            if (N == "sketchpoint") return swSelectType_e.swSelSKETCHPOINTS;
            if (N == "sketchsegment") return swSelectType_e.swSelSKETCHSEGS;
            if (N == "tableannotation") return swSelectType_e.swSelANNOTATIONTABLES;
            if (N == "titleblock") return swSelectType_e.swSelTITLEBLOCK;
            if (N == "titleblocktablefeature") return swSelectType_e.swSelTITLEBLOCKTABLEFEAT;
            if (N == "vertex") return swSelectType_e.swSelVERTICES;
            if (N == "view") return swSelectType_e.swSelDRAWINGVIEWS;
            if (N == "weldmentcutlistfeature") return swSelectType_e.swSelWELDMENTTABLEFEATS;
            if (N == "weldsymbol") return swSelectType_e.swSelWELDS;

            return swSelectType_e.swSelUNSUPPORTED;
        }

        //========================================================================================

        public static swDxfFormat_e swDxfDwg_Version
        {
            get { return (swDxfFormat_e)App.Sw.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion); }
            set { App.Sw.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion, (int)value); }
        }

        public static eDxfFormat DxfDwg_Version
        {
            get { return (eDxfFormat)App.Sw.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion); }
            set { App.Sw.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion, (int)value); }
        }

        public static Boolean DxfDwg_PolicesAutoCAD
        {
            get { return !Convert.ToBoolean(App.Sw.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputFonts)); }
            set { App.Sw.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputFonts, Convert.ToInt32(!value)); }
        }

        public static Boolean DxfDwg_StylesAutoCAD
        {
            get { return !Convert.ToBoolean(App.Sw.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputLineStyles)); }
            set { App.Sw.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputLineStyles, Convert.ToInt32(!value)); }
        }

        public static Boolean DxfDwg_SortieEchelle1
        {
            get { return Convert.ToBoolean(App.Sw.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputNoScale)); }
            set { App.Sw.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputNoScale, Convert.ToInt32(value)); }
        }

        public static Boolean DxfDwg_JoindreExtremites
        {
            get { return Convert.ToBoolean(App.Sw.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfEndPointMerge)); }
            set { App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfEndPointMerge, value); }
        }

        public static Double DxfDwg_JoindreExtremitesTolerance
        {
            get { return App.Sw.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swDxfMergingDistance); }
            set { App.Sw.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swDxfMergingDistance, value); }
        }

        public static Boolean DxfDwg_JoindreExtremitesHauteQualite
        {
            get { return Convert.ToBoolean(App.Sw.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDXFHighQualityExport)); }
            set { App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDXFHighQualityExport, value); }
        }

        public static Boolean DxfDwg_ExporterSplineEnPolyligne
        {
            get { return !Convert.ToBoolean(App.Sw.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportSplinesAsSplines)); }
            set { App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportSplinesAsSplines, !value); }
        }

        public static swDxfMultisheet_e DxfDwg_ExporterToutesLesFeuilles
        {
            get { return (swDxfMultisheet_e)App.Sw.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfMultiSheetOption); }
            set { App.Sw.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, (int)value); }
        }

        public static Boolean DxfDwg_ExporterFeuilleDansEspacePapier
        {
            get { return Convert.ToBoolean(App.Sw.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportAllSheetsToPaperSpace)); }
            set { App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportAllSheetsToPaperSpace, value); }
        }

        public static Boolean Pdf_ExporterEnCouleur
        {
            get { return Convert.ToBoolean(App.Sw.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportInColor)); }
            set { App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportInColor, value); }
        }

        public static Boolean Pdf_IncorporerLesPolices
        {
            get { return Convert.ToBoolean(App.Sw.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportEmbedFonts)); }
            set { App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportEmbedFonts, value); }
        }

        public static Boolean Pdf_ExporterEnHauteQualite
        {
            get { return Convert.ToBoolean(App.Sw.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportHighQuality)); }
            set { App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportHighQuality, value); }
        }

        public static Boolean Pdf_ImprimerEnTeteEtPiedDePage
        {
            get { return Convert.ToBoolean(App.Sw.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportPrintHeaderFooter)); }
            set { App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportPrintHeaderFooter, value); }
        }

        public static Boolean Pdf_UtiliserLesEpaisseursDeLigneDeImprimante
        {
            get { return Convert.ToBoolean(App.Sw.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportUseCurrentPrintLineWeights)); }
            set { App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportUseCurrentPrintLineWeights, value); }
        }

        //========================================================================================

        public static eTypeDoc eGetTypeDoc(this swDocumentTypes_e t)
        {
            switch (t)
            {
                case swDocumentTypes_e.swDocASSEMBLY:
                    return Outils.eTypeDoc.Assemblage;
                case swDocumentTypes_e.swDocDRAWING:
                    return Outils.eTypeDoc.Dessin;
                case swDocumentTypes_e.swDocLAYOUT:
                    break;
                case swDocumentTypes_e.swDocNONE:
                    break;
                case swDocumentTypes_e.swDocPART:
                    return Outils.eTypeDoc.Piece;
                case swDocumentTypes_e.swDocSDM:
                    break;
                default:
                    break;
            }

            return Outils.eTypeDoc.Inconnu;
        }

        public static swDocumentTypes_e eGetSwTypeDoc(this eTypeDoc t)
        {
            switch (t)
            {
                case Outils.eTypeDoc.Piece:
                    return swDocumentTypes_e.swDocPART;
                case Outils.eTypeDoc.Assemblage:
                    return swDocumentTypes_e.swDocASSEMBLY;
                case Outils.eTypeDoc.Dessin:
                    return swDocumentTypes_e.swDocDRAWING;
                case Outils.eTypeDoc.Inconnu:
                    return swDocumentTypes_e.swDocNONE;
                default:
                    break;
            }

            return swDocumentTypes_e.swDocNONE;
        }

        public static eTypeDoc TypeDoc(this ModelDoc2 mdl)
        {
            return eGetTypeDoc((swDocumentTypes_e)mdl.GetType());
        }

        public static eTypeDoc TypeDoc(this Component2 cp)
        {
            ModelDoc2 mdl = cp.GetModelDoc2();

            if (mdl == null) return Outils.eTypeDoc.Inconnu;

            return mdl.TypeDoc();
        }

        public static ModelDoc2 eModelDoc2(this PartDoc piece)
        {
            return piece as ModelDoc2;
        }

        public static ModelDoc2 eModelDoc2(this AssemblyDoc ass)
        {
            return ass as ModelDoc2;
        }

        public static ModelDoc2 eModelDoc2(this DrawingDoc dessin)
        {
            return dessin as ModelDoc2;
        }

        public static ModelDoc2 eModelDoc2(this Component2 cp)
        {
            return cp.GetModelDoc2() as ModelDoc2;
        }

        public static PartDoc ePartDoc(this ModelDoc2 mdl)
        {
            return mdl as PartDoc;
        }

        public static PartDoc ePartDoc(this Component2 cp)
        {
            return cp.GetModelDoc2() as PartDoc;
        }

        public static AssemblyDoc eAssemblyDoc(this ModelDoc2 mdl)
        {
            return mdl as AssemblyDoc;
        }

        public static AssemblyDoc eAssemblyDoc(this Component2 cp)
        {
            return cp.GetModelDoc2() as AssemblyDoc;
        }

        public static DrawingDoc eDrawingDoc(this ModelDoc2 mdl)
        {
            return mdl as DrawingDoc;
        }

        public static ModelDoc2 eModeleActif()
        {
            return App.Sw.ActiveDoc as ModelDoc2;
        }

        public static void eRenommerFonction(this Feature f, String nom)
        {
            if (f.IsNull()) return;

            String oldName = f.Name;
            int i = 0;

            while (f.Name == oldName)
            {
                f.Name = nom + ++i;

                if (i > 100) break;
            }
        }

        public static void eEditerLeComposant(this AssemblyDoc ass, Component2 composant)
        {
            ModelDoc2 mdl = ass.eModelDoc2();
            mdl.eEffacerSelection();
            composant.eSelect(false);
            int info = 0;
            ass.EditPart2(true, false, ref info);
            mdl.eEffacerSelection();
        }

        public static void eEditerAssemblage(this AssemblyDoc ass)
        {
            ass.EditAssembly();
        }

        public static void eActiver(this ModelDoc2 mdl, swRebuildOnActivation_e Reconstruire = swRebuildOnActivation_e.swUserDecision)
        {
            int Erreur = 0;
            App.Sw.ActivateDoc3(mdl.GetPathName(), true, (int)Reconstruire, Erreur);
            mdl.eZoomEtendu();
        }

        public static void eZoomEtendu(this ModelDoc2 mdl)
        {
            mdl.ViewZoomtofit2();
        }

        //========================================================================================

        public static String eNomSansExt(this Component2 cp)
        {
            return Path.GetFileNameWithoutExtension(cp.GetPathName());
        }

        public static String eNomSansExt(this ModelDoc2 mdl)
        {
            return Path.GetFileNameWithoutExtension(mdl.GetPathName());
        }

        public static String eNomAvecExt(this Component2 cp)
        {
            return Path.GetFileName(cp.GetPathName());
        }

        public static String eNomAvecExt(this ModelDoc2 mdl)
        {
            return Path.GetFileName(mdl.GetPathName());
        }

        /// <summary>
        /// Renvoi le chemin complet du dossier
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        public static String eDossier(this Component2 cp)
        {
            return Path.GetDirectoryName(cp.GetPathName());
        }

        /// <summary>
        /// Renvoi le chemin complet du dossier
        /// </summary>
        /// <param name="mdl"></param>
        /// <returns></returns>
        public static String eDossier(this ModelDoc2 mdl)
        {
            return Path.GetDirectoryName(mdl.GetPathName());
        }

        /// <summary>
        /// Renvoi le nom du dossier
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        public static String eNomDossier(this Component2 cp)
        {
            return new DirectoryInfo(cp.eDossier()).Name;
        }

        /// <summary>
        /// Renvoi le nom du dossier
        /// </summary>
        /// <param name="mdl"></param>
        /// <returns></returns>
        public static String eNomDossier(this ModelDoc2 mdl)
        {
            return new DirectoryInfo(mdl.eDossier()).Name;
        }

        public static Boolean eEstDansLeDossier(this ModelDoc2 mdl, ModelDoc2 mdlBase)
        {
            return mdl.GetPathName().Contains(mdlBase.eDossier());
        }

        public static Boolean eEstDansLeDossier(this Component2 cp, ModelDoc2 mdlBase)
        {
            return cp.GetPathName().Contains(mdlBase.eDossier());
        }

        public static ModelDoc2 eCreerDocument(String dossier, String nomDuDocument, eTypeDoc typeDeDocument, String gabarit = "")
        {
            String pCheminGabarit = "";

            nomDuDocument = string.Join("_", nomDuDocument.Split(Path.GetInvalidFileNameChars()));

            if (String.IsNullOrEmpty(gabarit))
            {
                switch (typeDeDocument)
                {
                    case Outils.eTypeDoc.Assemblage:
                        pCheminGabarit = App.Sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
                        break;
                    case Outils.eTypeDoc.Piece:
                        pCheminGabarit = App.Sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
                        break;
                    case Outils.eTypeDoc.Dessin:
                        pCheminGabarit = App.Sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
                        break;
                }
            }
            else
            {
                String[] pTabCheminsGabarit = App.Sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates).Split(';');

                foreach (String Chemin in pTabCheminsGabarit)
                {
                    pCheminGabarit = Chemin + @"\" + gabarit + typeDeDocument.GetEnumInfo<ExtGabarit>();
                    if (File.Exists(pCheminGabarit))
                        break;
                }
            }

            int Format = 0;
            Double Lg = 0;
            Double Ht = 0;

            if (typeDeDocument == Outils.eTypeDoc.Dessin)
            {
                Double[] pTab = App.Sw.GetTemplateSizes(pCheminGabarit);
                Format = (int)pTab[0];
                Lg = pTab[1];
                Ht = pTab[2];
            }

            int Erreur = 0, Warning = 0;

            ModelDoc2 Modele = App.Sw.NewDocument(pCheminGabarit, Format, Lg, Ht);
            Modele.Extension.SaveAs(Path.Combine(dossier, nomDuDocument + typeDeDocument.GetEnumInfo<ExtFichier>()),
                                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                        null,
                                        ref Erreur,
                                        ref Warning);

            return Modele;
        }

        /// <summary>
        /// Sauve le modele en Pdf 3D
        /// </summary>
        public static Boolean SauverEnPdf3D(this ModelDoc2 mdl, String Chemin)
        {
            if ((mdl.TypeDoc() == eTypeDoc.Assemblage) || (mdl.TypeDoc() == eTypeDoc.Piece))
            {
                int Erreur = 0, Warning = 0;
                ExportPdfData pExportPdfData = App.Sw.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);
                pExportPdfData.ExportAs3D = true;
                pExportPdfData.ViewPdfAfterSaving = false;
                return mdl.Extension.SaveAs(Chemin, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, pExportPdfData, ref Erreur, ref Warning);
            }

            return false;
        }

        //========================================================================================

        public static String eNomConfigActive(this ModelDoc2 mdl)
        {
            return App.Sw.GetActiveConfigurationName(mdl.GetPathName());
        }

        public static String eNomConfiguration(this Component2 cp)
        {
            String cfg = cp.ReferencedConfiguration;
            if (cp.IsRoot())
                cfg = cp.eModelDoc2().eNomConfigActive();

            return cfg;
        }

        public static List<String> eListeNomConfiguration(this ModelDoc2 mdl)
        {
            return ((String[])mdl.GetConfigurationNames()).ToList();
        }

        public static List<String> eListeNomConfiguration(this ModelDoc2 mdl, eTypeConfig filtre)
        {
            List<String> ListeNomsConfig = new List<String>();
            String[] ArrString = (String[])mdl.GetConfigurationNames();

            foreach (String Nom in ArrString)
            {
                Configuration Cfg = mdl.GetConfigurationByName(Nom);
                if (Cfg.eEst(filtre))
                    ListeNomsConfig.Add(Nom);
            }

            ListeNomsConfig.Sort(new WindowsStringComparer());

            return ListeNomsConfig;
        }

        public static List<String> eListeNomConfiguration(this ModelDoc2 mdl, Predicate<String> filtre)
        {
            List<String> ListeNomsConfig = new List<String>();
            String[] ArrString = (String[])mdl.GetConfigurationNames();

            foreach (String Nom in ArrString)
            {
                if (filtre(Nom))
                    ListeNomsConfig.Add(Nom);
            }

            ListeNomsConfig.Sort(new WindowsStringComparer());

            return ListeNomsConfig;
        }

        public static void eParcourirConfiguration(this ModelDoc2 mdl, Predicate<String> filtre)
        {
            String[] ArrString = (String[])mdl.GetConfigurationNames();

            foreach (String Nom in ArrString)
            {
                if (filtre(Nom))
                    return;
            }
        }

        public static void eParcourirConfiguration(this ModelDoc2 mdl, Predicate<Configuration> filtre)
        {
            eParcourirConfiguration(mdl,
                NomCfg =>
                {
                    Configuration Cfg = mdl.GetConfigurationByName(NomCfg);
                    return filtre(Cfg);
                }
                );
        }

        public static void eSetSupprimerNouvellesFonctions(this String config, Boolean val, ModelDoc2 mdl)
        {
            Configuration C = mdl.GetConfigurationByName(config);
            C.eSetSupprimerNouvellesFonctions(val, mdl);
        }

        public static void eSetSupprimerNouvellesFonctions(this Configuration config, Boolean val, ModelDoc2 mdl)
        {
            if (val)
            {
                mdl.EditConfiguration3(config.Name,
                                        config.Name,
                                        config.Comment,
                                        config.AlternateName,
                                        (int)swConfigurationOptions2_e.swConfigOption_InheritProperties + (int)swConfigurationOptions2_e.swConfigOption_SuppressByDefault
                                        );
            }
            else
            {
                mdl.EditConfiguration3(config.Name,
                                        config.Name,
                                        config.Comment,
                                        config.AlternateName,
                                        (int)swConfigurationOptions2_e.swConfigOption_InheritProperties
                                        );
            }
        }

        public static Boolean eEstConfigPliee(this Configuration Config)
        {
            return eEstConfigPliee(Config.Name);
        }

        public static Boolean eEstConfigPliee(this String NomConfig)
        {
            if (Regex.IsMatch(NomConfig, CONSTANTES.CONFIG_PLIEE_PATTERN))
                return true;

            return false;
        }

        public static Boolean eEstConfigDepliee(this Configuration Config)
        {
            return eEstConfigDepliee(Config.Name);
        }

        public static Boolean eEstConfigDepliee(this String NomConfig)
        {
            if (Regex.IsMatch(NomConfig, CONSTANTES.CONFIG_DEPLIEE_PATTERN))
                return true;

            return false;
        }

        public static eTypeConfig eTypeConfig(this Configuration config)
        {
            eTypeConfig T = 0;
            if (eEstConfigDepliee(config.Name))
                T = Outils.eTypeConfig.Depliee;
            else if (eEstConfigPliee(config.Name))
                T = Outils.eTypeConfig.Pliee;

            if (config.IsDerived() != false)
                T |= Outils.eTypeConfig.Derivee;
            else
                T |= Outils.eTypeConfig.Racine;

            if (!T.HasFlag(Outils.eTypeConfig.Depliee))
                T |= Outils.eTypeConfig.DeBase;

            return T;
        }

        public static Boolean eEst(this Configuration config, eTypeConfig T)
        {
            return (config.eTypeConfig() & T) != 0;
        }

        public static Configuration eConfigParent(this Configuration config)
        {
            Configuration pSwConfigurationParent = null;

            if (config.IsDerived() == true)
                pSwConfigurationParent = config.GetParent();

            return pSwConfigurationParent;
        }

        public static List<Configuration> eListeConfigs(this ModelDoc2 mdl, eTypeConfig typeConfig = Outils.eTypeConfig.Tous)
        {
            List<Configuration> Liste = new List<Configuration>();

            foreach (String pNomConfig in mdl.GetConfigurationNames())
            {
                Configuration pConfig = mdl.GetConfigurationByName(pNomConfig);

                if (pConfig.eEst(typeConfig))
                    Liste.Add(pConfig);
            }

            Liste.Sort((c1, c2) => { WindowsStringComparer sc = new WindowsStringComparer(); return sc.Compare(c1.Name, c2.Name); });

            return Liste;
        }

        public static String eEtatAffichageCourant(this ModelDoc2 mdl)
        {
            Configuration cfgBase = mdl.GetConfigurationByName(mdl.eNomConfigActive());

            return cfgBase.eEtatAffichageCourant();
        }

        public static String eEtatAffichageCourant(this Configuration cfg)
        {
            String NomCurrentDisplayState = "";
            String[] ld = cfg.GetDisplayStates();

            if (ld.IsRef() && (ld.Count() > 0))
                NomCurrentDisplayState = ld[0];

            return NomCurrentDisplayState;
        }

        public static Boolean eSupprimerConfigAvecEtatAff(this Configuration config, ModelDoc2 mdl)
        {
            String EtatAffichage = config.eEtatAffichageCourant();

            String[] pTabNomAff = config.GetDisplayStates();

            if (pTabNomAff.IsRef())
            {
                foreach (String pNomEtatAffichage in pTabNomAff)
                {
                    if (mdl.Extension.LinkedDisplayState)
                    {
                        config.DeleteDisplayState(pNomEtatAffichage);
                    }
                    else if (pNomEtatAffichage.StartsWith(CONSTANTES.ETAT_D_AFFICHAGE) || (pNomEtatAffichage != EtatAffichage))
                    {
                        config.DeleteDisplayState(pNomEtatAffichage);
                    }
                }
            }

            if (mdl.GetConfigurationCount() > 1)
            {
                Boolean r = mdl.DeleteConfiguration2(config.Name);
                mdl.Extension.PurgeDisplayState();

                return r;
            }

            return false;
        }

        public static void eRenommerEtatAffichage(this Configuration Config, Boolean Ecraser = false)
        {
            String NomCfg = Config.Name;

            int Index = 1;

            if (Config.GetDisplayStatesCount() > 0)
            {
                foreach (String pNomEtatAffichage in Config.GetDisplayStates())
                {
                    if (pNomEtatAffichage.StartsWith(CONSTANTES.ETAT_D_AFFICHAGE) || Ecraser)
                    {
                        String NomTmp = NomCfg;
                        if (!Config.RenameDisplayState(pNomEtatAffichage, NomTmp))
                        {
                            NomTmp = NomCfg + "_" + Index.ToString();
                            Index++;
                        }
                    }
                }
            }
        }

        public static void eSupprimerEtatAffichage(this Configuration Config, ModelDoc2 mdl, String EtatAffichageBase = "", Boolean Tous = false)
        {
            String EtatAffichageCourant = "";

            if (String.IsNullOrWhiteSpace(EtatAffichageBase))
                EtatAffichageCourant = mdl.eEtatAffichageCourant();
            else
                EtatAffichageCourant = EtatAffichageBase;

            mdl.Extension.PurgeDisplayState();

            String[] ld = Config.GetDisplayStates();
            Config.ApplyDisplayState(EtatAffichageCourant);

            if (ld.IsRef() && (ld.Count() > 0))
            {
                foreach (String NomEtatAffichage in ld)
                {
                    if ((Tous || NomEtatAffichage.StartsWith(CONSTANTES.ETAT_D_AFFICHAGE)) && (NomEtatAffichage != EtatAffichageCourant))
                        Config.DeleteDisplayState(NomEtatAffichage);
                }
            }

            mdl.Extension.PurgeDisplayState();
        }

        /// <summary>
        /// Ajoute une configuration
        /// Si les Etats d'affichage sont liés, ils seront renommés
        /// Sinon, ils seront supprimés
        /// </summary>
        /// <param name="mdl"></param>
        /// <param name="name"></param>
        /// <param name="comment"></param>
        /// <param name="alternateName"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public static Configuration eAddConfiguration(this ModelDoc2 mdl, string name, string comment, string alternateName, int options)
        {
            Boolean IsLinked = mdl.Extension.LinkedDisplayState;
            String NomCurrentDisplayState = mdl.eEtatAffichageCourant();

            mdl.AddConfiguration3(name, comment, alternateName, options);
            Configuration cfg = mdl.GetConfigurationByName(name);

            if (IsLinked)
                cfg.eRenommerEtatAffichage();
            else
            {
                mdl.Extension.LinkedDisplayState = IsLinked;
                cfg.eSupprimerEtatAffichage(mdl, NomCurrentDisplayState);
            }

            return cfg;
        }

        public static Component2 eComposantRacine(this ModelDoc2 mdl, Boolean Maj = true)
        {
            return mdl.ConfigurationManager.ActiveConfiguration.GetRootComponent3(Maj);
        }

        /// <summary>
        /// Recherche de façon recursive le premier composant validant les critères de predicate
        /// </summary>
        /// <param name="mdl"></param>
        /// <param name="filtreTest">Filtre</param>
        /// <returns></returns>
        public static Component2 eRecChercherComposant(this ModelDoc2 mdl, Predicate<Component2> filtreTest, Predicate<Component2> filtreRec = null)
        {
            return mdl.eComposantRacine().eRecChercherComposant(filtreTest, filtreRec);
        }

        /// <summary>
        /// Recherche de façon recursive le premier composant validant les critères de predicate
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="filtreTest">Filtre</param>
        /// <returns></returns>
        public static Component2 eRecChercherComposant(this Component2 cp, Predicate<Component2> filtreTest, Predicate<Component2> filtreRec = null)
        {
            if (cp.TypeDoc() == Outils.eTypeDoc.Assemblage)
            {
                List<Component2> Liste = new List<Component2>();
                cp.eRecListeComposantBase(ref Liste, filtreTest, filtreRec, true);

                if (Liste.Count > 0)
                    return Liste[0];
            }

            return null;
        }

        /// <summary>
        /// Liste de façon recursive les composants
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="filtreAdd">Filtre</param>
        /// <param name="trier"></param>
        /// <returns></returns>
        public static List<Component2> eRecListeComposant(this ModelDoc2 mdl, Predicate<Component2> filtreAdd = null, Predicate<Component2> filtreRec = null, Boolean trier = false)
        {
            return mdl.eComposantRacine().eRecListeComposant(filtreAdd, filtreRec, trier);
        }

        /// <summary>
        /// Liste de façon recursive les composants
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="filtreAdd">Filtre</param>
        /// <param name="trier"></param>
        /// <returns></returns>
        public static List<Component2> eRecListeComposant(this Component2 cp, Predicate<Component2> filtreAdd = null, Predicate<Component2> filtreRec = null, Boolean trier = false)
        {
            List<Component2> Liste = new List<Component2>();

            if (cp.TypeDoc() == Outils.eTypeDoc.Assemblage)
            {
                cp.eRecListeComposantBase(ref Liste, filtreAdd, filtreRec, false);

                if (trier)
                    Liste.Sort((c1, c2) => { WindowsStringComparer sc = new WindowsStringComparer(); return sc.Compare(c1.eKeyAvecConfig(), c2.eKeyAvecConfig()); });
            }

            return Liste;
        }

        /// <summary>
        /// Liste de façon recursive les composants
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="Liste"></param>
        /// <param name="filtreAdd"></param>
        /// <param name="premiereOccurence"></param>
        /// <returns>Renvoi true si la fonction a trouvée la première occurence, sinon false</returns>
        public static Boolean eRecListeComposantBase(this Component2 cp, ref List<Component2> Liste, Predicate<Component2> filtreAdd, Predicate<Component2> filtreRec, Boolean premiereOccurence)
        {
            Object[] ChildComp = (Object[])cp.GetChildren();

            if (ChildComp.IsNull())
                return false;

            foreach (Component2 Cp in ChildComp)
            {
                if (filtreAdd.IsNull() || filtreAdd(Cp))
                {
                    Liste.Add(Cp);

                    if (premiereOccurence)
                        return true;
                }

                if (!Cp.IsSuppressed() && (filtreRec.IsNull() || filtreRec(Cp)))
                {
                    // Si on traverse les composants et que la recherche renvoi true,
                    // on a trouvé la première occurence donc on remonte true
                    if (Cp.eRecListeComposantBase(ref Liste, filtreAdd, filtreRec, premiereOccurence))
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Parcours de façon recursive les composants
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="filtreApp"></param>
        public static void eRecParcourirComposantBase(this Component2 cp, Action<Component2> filtreApp, Predicate<Component2> filtreRec = null)
        {
            Object[] ChildComp = (Object[])cp.GetChildren();

            if (ChildComp.IsNull())
                return;

            foreach (Component2 Cp in ChildComp)
            {
                if (filtreApp.IsRef())
                    filtreApp(Cp);

                if (!Cp.IsSuppressed() && Cp.TypeDoc() == eTypeDoc.Assemblage && (filtreRec.IsNull() || filtreRec(Cp)))
                    Cp.eRecParcourirComposantBase(filtreApp, filtreRec);
            }
        }

        /// <summary>
        /// Parcoure les composants de façon recursive et s'arrete si filtre renvoi true
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="filtreTest"></param>
        /// <param name="reccurence"></param>
        /// <returns></returns>
        public static Boolean eRecParcourirComposants(this Component2 cp, Predicate<Component2> filtreTest, Predicate<Component2> filtreRec = null)
        {
            Object[] ChildComp = (Object[])cp.GetChildren();

            if ((cp.TypeDoc() == eTypeDoc.Piece) || ChildComp.IsNull())
                return false;

            foreach (Component2 Cp in ChildComp)
            {
                if (filtreTest(Cp)) return true;

                if (!Cp.IsSuppressed() && (filtreRec.IsNull() || filtreRec(Cp)))
                {
                    // Si on traverse les composants et que la recherche renvoi true,
                    // on a trouvé la première occurence donc on remonte true
                    if (Cp.eRecParcourirComposants(filtreTest, filtreRec))
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Parcoure les composants de façon recursive et s'arrete si filtre renvoi true
        /// </summary>
        /// <param name="mdl"></param>
        /// <param name="filtreTest"></param>
        /// <param name="reccurence"></param>
        /// <returns></returns>
        public static Boolean eRecParcourirComposants(this ModelDoc2 mdl, Predicate<Component2> filtreTest, Predicate<Component2> filtreRec = null)
        {
            return eRecParcourirComposants(mdl.eComposantRacine(), filtreTest, filtreRec);
        }

        public static void eParcourirComposants(this ModelDoc2 mdl, Predicate<Component2> fonction)
        {
            if (mdl.TypeDoc() == eTypeDoc.Assemblage)
            {
                Object[] ChildComp = (Object[])mdl.eAssemblyDoc().GetComponents(false);
                foreach (Component2 Cp in ChildComp)
                    if (fonction(Cp)) return;
            }
        }

        public static List<Component2> eListeComposants(this ModelDoc2 mdl)
        {
            List<Component2> Liste = new List<Component2>();

            if (mdl.TypeDoc() == eTypeDoc.Assemblage)
            {
                Object[] ChildComp = (Object[])mdl.eAssemblyDoc().GetComponents(false);
                foreach (Component2 Cp in ChildComp)
                    Liste.Add(Cp);
            }

            return Liste;
        }

        /// <summary>
        /// Retourne la liste des composants parent sans le composant racine, sans le composant de base
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        public static List<Component2> eListeComposantParent(this Component2 cp)
        {
            List<Component2> Liste = new List<Component2>();
            Component2 C = cp.GetParent();

            while ((C.IsRef()) && !C.IsRoot())
            {
                Liste.Add(C);
                C = C.GetParent();
            }

            return Liste;
        }

        /// <summary>
        /// Composant racine
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        public static Component2 eComposantParentRacine(this Component2 cp)
        {
            Component2 C = cp;

            while (!C.IsRoot())
            {
                C = C.GetParent();
            }

            return C;
        }

        /// <summary>
        /// Renvoi le niveau du composant dans la hiérarchie, 0 etant le composant racine
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        public static int eNiveauComposant(this Component2 cp)
        {
            int n = 0;
            if (cp.IsRoot())
                return 0;

            n++;
            Component2 C = cp.GetParent();

            while ((C.IsRef()) && !C.IsRoot())
            {
                n++;
                C = C.GetParent();
            }

            return n;
        }

        public static SortedDictionary<ModelDoc2, SortedDictionary<String, int>> eDenombrerComposant(this Component2 cp, Predicate<Component2> filtre = null, Predicate<Component2> filtreRec = null)
        {
            SortedDictionary<ModelDoc2, SortedDictionary<String, int>> Dic = new SortedDictionary<ModelDoc2, SortedDictionary<string, int>>(new CompareModelDoc2());

            if (cp.TypeDoc() == eTypeDoc.Assemblage)
                cp.eRecParcourirComposantBase(
                    c =>
                    {
                        if (!c.IsSuppressed() && filtre(c))
                        {
                            var mdl = c.eModelDoc2();
                            var cfg = c.eNomConfiguration();

                            if (Dic.ContainsKey(mdl))
                            {
                                var dicCfg = Dic[mdl];
                                dicCfg.AddIfNotExistOrPlus(cfg);
                            }
                            else
                            {
                                SortedDictionary<string, int> dicCfg = new SortedDictionary<string, int>(new WindowsStringComparer());
                                dicCfg.Add(cfg, 1);
                                Dic.Add(mdl, dicCfg);
                            }
                        }
                    }
                    ,
                    filtreRec
                    );

            return Dic;
        }

        //========================================================================================

        public static String eNom(this BodyFolder dossier) { return dossier.GetFeature().Name; }

        /// <summary>
        /// Retourne si le dossier est exclu de la nomenclature.
        /// </summary>
        public static Boolean eEstExclu(this BodyFolder dossier)
        {
            return Convert.ToBoolean(dossier.GetFeature().ExcludeFromCutList);
        }

        /// <summary>
        /// Retourne le type de corps du dossier.
        /// </summary>
        public static eTypeCorps eTypeDeDossier(this BodyFolder dossier)
        {
            if (dossier.eEstUnDossierDeToles())
                return eTypeCorps.Tole;

            if (dossier.eEstUnDossierDeBarres())
                return eTypeCorps.Barre;

            if (dossier.GetBodyCount() > 0)
            {
                Body2 corps = dossier.GetBodies()[0];
                return corps.eTypeDeCorps();
            }

            return eTypeCorps.Autre;
        }

        /// <summary>
        /// Test si le dossier contient des barres en recherchant la propriete "LENGTH"
        /// </summary>
        /// <param name="dossier"></param>
        /// <returns></returns>
        public static Boolean eEstUnDossierDeBarres(this BodyFolder dossier)
        {
            CustomPropertyManager pPropMgr = dossier.GetFeature().CustomPropertyManager;

            foreach (String iNom in pPropMgr.GetNames())
            {
                String pVal, pResult;
                Boolean Resolved;

                pPropMgr.Get5(iNom, false, out pVal, out pResult, out Resolved);

                if (Regex.IsMatch(pVal, "^\"LENGTH@@@"))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Test si le dossier contient des toles en recherchant la propriete "SW-Longueur du flanc de tôle"
        /// </summary>
        /// <param name="dossier"></param>
        /// <returns></returns>
        public static Boolean eEstUnDossierDeToles(this BodyFolder dossier)
        {
            CustomPropertyManager pPropMgr = dossier.GetFeature().CustomPropertyManager;

            foreach (String nom in pPropMgr.GetNames())
            {
                String pVal, pResult;
                Boolean Resolved;

                pPropMgr.Get5(nom, false, out pVal, out pResult, out Resolved);

                if (Regex.IsMatch(pVal, "^\"SW-Longueur du flanc de tôle@@@"))
                {
                    return true;
                }
            }

            return false;
        }

        public static List<Body2> eListeDesCorps(this BodyFolder dossier)
        {
            List<Body2> Liste = new List<Body2>();

            if (dossier.GetBodyCount() > 0)
            {
                foreach (Body2 corps in dossier.GetBodies())
                    Liste.Add(corps);
            }

            return Liste;
        }

        public static int eNbCorps(this BodyFolder dossier)
        {
            if (dossier.IsNull()) return 0;

            return dossier.GetBodyCount();
        }

        public static Body2 ePremierCorps(this BodyFolder dossier)
        {
            if (dossier.GetBodyCount() == 0) return null;

            object[] Corps = dossier.GetBodies();

            return Corps[0] as Body2;
        }

        public static Boolean eInsererListeDesPiecesSoudees(this PartDoc piece)
        {
            if (piece.eDossierListeDesPiecesSoudees().IsRef()) return false;

            piece.eModelDoc2().FeatureManager.InsertWeldmentFeature();

            return true;
        }

        public static void eMajListeDesPiecesSoudees(this Component2 composant)
        {
            if (composant.TypeDoc() != eTypeDoc.Piece) return;

            if (composant.IsRoot())
                composant.ePartDoc().eDossierListeDesPiecesSoudees().eMajListeDesPiecesSoudees();
            else
                composant.eDossierListeDesPiecesSoudees().eMajListeDesPiecesSoudees();
        }

        public static void eMajListeDesPiecesSoudees(this PartDoc piece)
        {
            piece.eDossierListeDesPiecesSoudees().eMajListeDesPiecesSoudees();
        }

        private static void eMajListeDesPiecesSoudees(this Feature fonction)
        {
            if (fonction.IsNull()) return;

            BodyFolder dossier = fonction.GetSpecificFeature2();

            if (dossier.IsNull()) return;

            dossier.UpdateCutList();
        }

        private static Feature eDossierListeDesPiecesSoudees(this Feature fonction)
        {
            Feature pFonction = fonction;

            while (pFonction.IsRef())
            {
                if (pFonction.GetTypeName2() == "SolidBodyFolder")
                    return pFonction;

                pFonction = pFonction.GetNextFeature();
            }

            return null;
        }

        public static Feature eDossierListeDesPiecesSoudees(this PartDoc piece)
        {
            Feature pFonction = piece.FirstFeature();

            return pFonction.eDossierListeDesPiecesSoudees();
        }

        public static Feature eDossierListeDesPiecesSoudees(this Component2 composant)
        {
            if (composant.TypeDoc() != eTypeDoc.Piece)
            {
                Log.Methode("Sw", (Object)("Le composant n'est pas une pièce"));
                return null;
            }

            Feature pFonction = composant.FirstFeature();

            if (composant.IsRoot())
                pFonction = composant.ePartDoc().FirstFeature();

            return pFonction.eDossierListeDesPiecesSoudees();
        }

        public static void eParcourirFonctionsDePiecesSoudees(this Feature Fonction, Action<Feature> action)
        {
            Feature f = Fonction;

            if (f.IsRef())
            {
                f = f.GetFirstSubFeature();

                while (f.IsRef())
                {
                    action(f);

                    f.eParcourirFonctionsDePiecesSoudees(action);

                    f = f.GetNextSubFeature();
                }
            }
        }

        public static List<BodyFolder> eListeDesDossiersDePiecesSoudees(this Feature Fonction, Predicate<BodyFolder> filtre = null)
        {
            List<BodyFolder> Liste = new List<BodyFolder>();

            Fonction.eParcourirFonctionsDePiecesSoudees(
                f =>
                {
                    if (f.GetTypeName2() == FeatureType.swTnCutListFolder)
                    {
                        BodyFolder dossier = f.GetSpecificFeature2();
                        if (dossier.IsRef() && (dossier.GetBodyCount() > 0) && (filtre.IsNull() || filtre(dossier)))
                            Liste.Add(dossier);
                    }
                }
                );

            return Liste;
        }

        public static List<BodyFolder> eListeDesDossiersDePiecesSoudees(this PartDoc piece, Predicate<BodyFolder> filtre = null)
        {
            return piece.eDossierListeDesPiecesSoudees().eListeDesDossiersDePiecesSoudees(filtre);
        }

        /// <summary>
        /// A utiliser avec attention, le contenu des dossiers n'est pas toujours à jour
        /// </summary>
        /// <param name="composant"></param>
        /// <param name="filtre"></param>
        /// <returns></returns>
        public static List<BodyFolder> eListeDesDossiersDePiecesSoudees(this Component2 composant, Predicate<BodyFolder> filtre = null)
        {
            return composant.eDossierListeDesPiecesSoudees().eListeDesDossiersDePiecesSoudees(filtre);
        }

        public static List<Feature> eListeDesFonctionsDePiecesSoudees(this Feature Fonction, Predicate<Feature> filtre = null)
        {
            List<Feature> Liste = new List<Feature>();

            Fonction.eParcourirFonctionsDePiecesSoudees(
                f =>
                {
                    if ((f.GetTypeName2() == FeatureType.swTnCutListFolder) && (filtre.IsNull() || filtre(f)))
                        Liste.Add(f);
                }
                );

            return Liste;
        }

        public static List<Feature> eListeDesFonctionsDePiecesSoudees(this PartDoc piece, Predicate<Feature> filtre = null)
        {
            return piece.eDossierListeDesPiecesSoudees().eListeDesFonctionsDePiecesSoudees(filtre);
        }

        public static ListPID<Feature> eListePIDdesFonctionsDePiecesSoudees(this Feature Fonction, ModelDoc2 mdl, Predicate<Feature> filtre = null)
        {
            ListPID<Feature> Liste = new ListPID<Feature>(mdl);

            Fonction.eParcourirFonctionsDePiecesSoudees(
                f =>
                {
                    if ((f.GetTypeName2() == FeatureType.swTnCutListFolder) && (filtre.IsNull() || filtre(f)))
                        Liste.Add(f);
                }
                );

            return Liste;
        }

        public static ListPID<Feature> eListePIDdesFonctionsDePiecesSoudees(this PartDoc piece, Predicate<Feature> filtre = null)
        {
            return piece.eDossierListeDesPiecesSoudees().eListePIDdesFonctionsDePiecesSoudees(piece.eModelDoc2(), filtre);
        }

        public static ListPID<Feature> eListePIDdesFonctionsDeSousEnsembleDePiecesSoudees(this Feature Fonction, ModelDoc2 mdl)
        {
            ListPID<Feature> Liste = new ListPID<Feature>(mdl);

            Fonction.eParcourirFonctionsDePiecesSoudees(
                f =>
                {
                    if (f.GetTypeName2() == FeatureType.swTnSubWeldFolder)
                        Liste.Add(f);
                }
                );

            return Liste;
        }

        public static SolidWorks.Interop.sldworks.Attribute eRecupererAttribut(this ModelDoc2 mdl, String nomAtt)
        {
            SolidWorks.Interop.sldworks.Attribute Att = null;
            Feature F = mdl.eChercherFonction(f => { return f.Name == nomAtt; });
            if (F.IsRef())
                Att = F.GetSpecificFeature2();

            return Att;
        }

        public static CustomPropertyManager eGestProp(this ModelDoc2 mdl, String config = "")
        {
            return mdl.Extension.CustomPropertyManager[config];
        }

        public static CustomPropertyManager eGestProp(this BodyFolder dossier)
        {
            return dossier.GetFeature().CustomPropertyManager;
        }

        public static Boolean ePropExiste(this CustomPropertyManager pm, String nomPropriete, Boolean CaseSensitive = false)
        {
            string[] CustomPropNames = (string[])pm.GetNames();

            String compare = CaseSensitive ? nomPropriete : nomPropriete.RemoveDiacritics().ToUpperInvariant();

            if (CustomPropNames.TabIsRef_LgthNotNull())
            {
                foreach (String n in CustomPropNames)
                {
                    String test = CaseSensitive ? n : n.RemoveDiacritics().ToUpperInvariant();
                    if (compare == test)
                        return true;
                }
            }

            return false;
        }

        public static Boolean ePropExiste(this ModelDoc2 mdl, String nomPropriete, String Config = "")
        {
            return mdl.eGestProp(Config).ePropExiste(nomPropriete);
        }

        /// <summary>
        /// Test si une propriete existe dans la configuration référencée
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="nomPropriete"></param>
        /// <returns></returns>
        public static Boolean ePropExiste(this Component2 cp, String nomPropriete)
        {
            ModelDoc2 mdl = cp.GetModelDoc2();
            return mdl.ePropExiste(nomPropriete, cp.eNomConfiguration());
        }

        public static Boolean ePropExiste(this BodyFolder dossier, String nomPropriete, Boolean CaseSensitive = false)
        {
            CustomPropertyManager pm = dossier.eGestProp();
            return pm.ePropExiste(nomPropriete, CaseSensitive);
        }

        public static String eProp(this ModelDoc2 mdl, String nomPropriete, String nomConfig = "")
        {
            String val, result = ""; Boolean wasResolved, link;

            CustomPropertyManager PM = mdl.eGestProp(nomConfig);
            if (PM.IsRef())
                PM.Get6(nomPropriete, false, out val, out result, out wasResolved, out link);

            return result;
        }

        public static swCustomInfoDeleteResult_e ePropSuppr(this ModelDoc2 mdl, String nomPropriete, String nomConfig = "")
        {
            swCustomInfoDeleteResult_e result = swCustomInfoDeleteResult_e.swCustomInfoDeleteResult_OK;

            CustomPropertyManager PM = mdl.eGestProp(nomConfig);
            if (PM.IsRef())
                result = (swCustomInfoDeleteResult_e)PM.Delete2(nomPropriete);

            return result;
        }

        /// <summary>
        /// Renvoi la valeur de la propriete de la configuration référencée
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="nomPropriete"></param>
        /// <returns></returns>
        public static String eProp(this Component2 cp, String nomPropriete)
        {
            ModelDoc2 mdl = cp.GetModelDoc2();
            return mdl.eProp(nomPropriete, cp.eNomConfiguration());
        }

        public static String eProp(this BodyFolder dossier, String nomPropriete)
        {
            CustomPropertyManager PM = dossier.eGestProp();
            String val, result = ""; Boolean wasResolved;
            PM.Get5(nomPropriete, false, out val, out result, out wasResolved);

            return result;
        }

        public static swCustomInfoAddResult_e ePropAdd(this CustomPropertyManager pm, String nomPropriete, Object val, swCustomPropertyAddOption_e action = swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
        {
            swCustomInfoAddResult_e r = (swCustomInfoAddResult_e)pm.Add3(nomPropriete, (int)swCustomInfoType_e.swCustomInfoText, val.ToString(), (int)action);

            return r;
        }

        public static swCustomInfoAddResult_e ePropAdd(this ModelDoc2 mdl, String nomPropriete, Object val, String nomConfig = "", swCustomPropertyAddOption_e action = swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
        {
            CustomPropertyManager PM = mdl.eGestProp(nomConfig);

            swCustomInfoAddResult_e r = (swCustomInfoAddResult_e)PM.ePropAdd(nomPropriete, val, action);

            return r;
        }

        public static void ePropSet(this CustomPropertyManager pm, String nomPropriete, Object val)
        {
            pm.Set2(nomPropriete, val.ToString());
        }

        public static Dictionary<String, String> eListProp(this BodyFolder bf)
        {
            return bf.eGestProp().eListProp();
        }

        public static Dictionary<String, String> eListProp(this CustomPropertyManager pm)
        {
            Dictionary<String, String> dic = new Dictionary<String, String>();

            object vPropNames = null;
            object vPropTypes = null;
            object vPropValues = null;
            object resolved = null;

            int nbProp = pm.GetAll2(ref vPropNames, ref vPropTypes, ref vPropValues, ref resolved);

            string[] Names = (string[])vPropNames;
            string[] Values = (string[])vPropValues;

            if (Names.TabIsRef_LgthNotNull())
            {
                for (int i = 0; i < nbProp; i++)
                    dic.Add(Names[i], Values[i]);
            }

            return dic;
        }

        public static Dictionary<String, String> eListProp(this ModelDoc2 mdl, String nomConfig = "")
        {
            return mdl.eGestProp(nomConfig).eListProp();
        }

        public static String eRefFichier(this ModelDoc2 mdl)
        {
            var Ref = String.Format("{0}-{1}", mdl.eProp(CONSTANTES.PROPRIETE_NOCLIENT), mdl.eProp(CONSTANTES.PROPRIETE_NOCOMMANDE)).Trim();

            return (Ref.StartsWith("-") || Ref.EndsWith("-")) ? "" : Ref;
        }

        //========================================================================================

        public static List<Feature> eListeFonctions(this ModelDoc2 mdl, Predicate<Feature> filtre, Boolean sousFonction = false)
        {
            Feature F = mdl.FirstFeature();
            return F.eListeFonctionsBase(filtre, sousFonction, false);
        }

        public static List<Feature> eListeFonctions(this Component2 cp, Predicate<Feature> filtre, Boolean sousFonction = false)
        {
            Feature F = cp.FirstFeature();
            return F.eListeFonctionsBase(filtre, sousFonction, false);
        }

        public static Feature eChercherFonction(this ModelDoc2 mdl, Predicate<Feature> filtre, Boolean sousFonction = false)
        {
            Feature F = (Feature)mdl.FirstFeature();

            List<Feature> ListeF = F.eListeFonctionsBase(filtre, sousFonction, true);

            Feature newF = null;

            if (ListeF.Count > 0)
                newF = ListeF[0];

            return newF;
        }

        public static Feature eChercherFonction(this Component2 cp, Predicate<Feature> filtre, Boolean sousFonction = false)
        {
            Feature F = (Feature)cp.FirstFeature();

            List<Feature> ListeF = F.eListeFonctionsBase(filtre, sousFonction, true);

            Feature newF = null;

            if (ListeF.Count > 0)
                newF = ListeF[0];

            return newF;
        }

        private static List<Feature> eListeFonctionsBase(this Feature f, Predicate<Feature> filtre, Boolean sousFonction, Boolean premiereOccurence)
        {
            List<Feature> pListeFonctions = new List<Feature>();

            Feature pSwFonction = f;

            while (pSwFonction != null)
            {
                if ((filtre.IsNull() || filtre(pSwFonction)) && !pListeFonctions.Contains(pSwFonction))
                {
                    pListeFonctions.Add(pSwFonction);
                    if (premiereOccurence)
                        return pListeFonctions;

                    if (sousFonction)
                    {
                        Feature pSwSousFonction = pSwFonction.GetFirstSubFeature();

                        while (pSwSousFonction != null)
                        {
                            if ((filtre.IsNull() || filtre(pSwSousFonction)) && !pListeFonctions.Contains(pSwFonction))
                            {
                                pListeFonctions.Add(pSwSousFonction);

                                if (premiereOccurence)
                                    return pListeFonctions;
                            }

                            pSwSousFonction = pSwSousFonction.GetNextSubFeature();
                        }
                    }
                }

                pSwFonction = pSwFonction.GetNextFeature();
            }

            return pListeFonctions;
        }

        public static void eParcourirFonctions(this Component2 cp, Predicate<Feature> filtre, Boolean sousFonction)
        {
            Feature f = cp.FirstFeature();

            f.eParcourirFonctions(filtre, sousFonction);
        }

        public static void eParcourirFonctions(this ModelDoc2 mdl, Predicate<Feature> filtre, Boolean sousFonction)
        {
            Feature f = mdl.FirstFeature();

            f.eParcourirFonctions(filtre, sousFonction);
        }

        private static void eParcourirFonctions(this Feature f, Predicate<Feature> filtre, Boolean sousFonction)
        {
            Feature pSwFonction = f;

            while (pSwFonction != null)
            {
                if (filtre(pSwFonction))
                    return;

                if (sousFonction && f.eParcourirSousFonction(filtre))
                    return;

                pSwFonction = pSwFonction.GetNextFeature();
            }
        }

        public static List<Feature> eListeSousFonction(this Feature f)
        {
            return f.eListeSousFonction(null);
        }

        public static List<Feature> eListeSousFonction(this Feature f, Predicate<Feature> filtre)
        {
            List<Feature> pListeFonctions = new List<Feature>();

            Feature pSwFonction = f.GetFirstSubFeature();

            while (pSwFonction != null)
            {
                if ((filtre.IsNull() || filtre(pSwFonction)) && !pListeFonctions.Contains(pSwFonction))
                    pListeFonctions.Add(pSwFonction);

                pSwFonction = pSwFonction.GetNextSubFeature();
            }

            return pListeFonctions;
        }

        public static Boolean eParcourirSousFonction(this Feature f, Predicate<Feature> filtre)
        {
            Feature pSwFonction = f.GetFirstSubFeature();

            while (pSwFonction != null)
            {
                if (filtre(pSwFonction))
                    return true;

                pSwFonction = pSwFonction.GetNextSubFeature();
            }

            return false;
        }

        public static Feature eFonctionParent(this Feature f)
        {
            Feature Parent = f.GetOwnerFeature();
            if (Parent.IsRef())
                return Parent;

            return null;
        }

        public static void eParcourirFonctions(this Body2 corps, Predicate<Feature> filtre, Boolean sousFonction)
        {
            foreach (Feature f in corps.GetFeatures())
            {
                if (filtre(f))
                    return;

                if (sousFonction && f.eParcourirSousFonction(filtre))
                    return;
            }
        }

        public static List<Feature> eListeFonctionsBase(this Body2 corps, Predicate<Feature> filtre, Boolean sousFonction, Boolean premiereOccurence)
        {
            List<Feature> pListeFonctions = new List<Feature>();

            foreach (Feature pSwFonction in corps.GetFeatures())
            {
                if ((filtre.IsNull() || filtre(pSwFonction)) && !pListeFonctions.Contains(pSwFonction))
                {
                    pListeFonctions.Add(pSwFonction);
                    if (premiereOccurence)
                        return pListeFonctions;

                    if (sousFonction)
                    {
                        Feature pSwSousFonction = pSwFonction.GetFirstSubFeature();

                        while (pSwSousFonction != null)
                        {
                            if ((filtre.IsNull() || filtre(pSwSousFonction)) && !pListeFonctions.Contains(pSwFonction))
                            {
                                pListeFonctions.Add(pSwSousFonction);

                                if (premiereOccurence)
                                    return pListeFonctions;
                            }

                            pSwSousFonction = pSwSousFonction.GetNextSubFeature();
                        }
                    }
                }
            }

            return pListeFonctions;
        }

        public static List<Feature> eListeFonctions(this Body2 corps, Predicate<Feature> filtre, Boolean sousFonction = false)
        {
            return corps.eListeFonctionsBase(filtre, sousFonction, false);
        }

        public static Feature eChercherFonction(this Body2 corps, Predicate<Feature> filtre, Boolean sousFonction = false)
        {
            List<Feature> ListeF = corps.eListeFonctionsBase(filtre, sousFonction, true);

            Feature newF = null;

            if (ListeF.Count > 0)
                newF = ListeF[0];

            return newF;
        }

        //========================================================================================

        public static Boolean eModifierEtat(this Feature f, swFeatureSuppressionAction_e etat)
        {

            Feature Fonction = f;

            if (f.GetTypeName2() == "SketchBlockInst")
                Fonction = f.eFonctionParent();

            return Fonction.SetSuppression2((int)etat,
                                            (int)swInConfigurationOpts_e.swThisConfiguration,
                                            null);

        }

        public static Boolean eModifierEtat(this Feature f, swFeatureSuppressionAction_e etat, String config)
        {

            Feature Fonction = f;

            if (f.GetTypeName2() == "SketchBlockInst")
                Fonction = f.eFonctionParent();

            if (!String.IsNullOrWhiteSpace(config))
            {
                return Fonction.eModifierEtat(etat, new List<string>() { config });
            }
            else
            {
                return Fonction.eModifierEtat(etat);
            }

        }

        public static Boolean eModifierEtat(this Feature f, swFeatureSuppressionAction_e etat, List<String> listeConfig)
        {

            Feature Fonction = f;

            if (f.GetTypeName2() == "SketchBlockInst")
                Fonction = f.eFonctionParent();

            String[] pTabConfig = listeConfig.ToArray();

            return Fonction.SetSuppression2((int)etat,
                                        (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                                        pTabConfig);

        }

        public static Boolean eEstDesactive(this Feature f)
        {
            Boolean[] pArrayResult = (Boolean[])f.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, null);
            if ((pArrayResult.IsRef()) && (pArrayResult[0] == true))
                return true;

            return false;
        }

        public static Boolean eModifierEtatComposant(this AssemblyDoc ass, swComponentSuppressionState_e etat)
        {
            return ass.SetComponentSuppression((int)etat);
        }

        public static void eFixer(this Component2 cp, AssemblyDoc ass)
        {
            cp.eSelectById(ass.eModelDoc2());
            ass.FixComponent();
        }

        public static void eLiberer(this Component2 cp, AssemblyDoc ass)
        {
            cp.eSelectById(ass.eModelDoc2());
            ass.UnfixComponent();
        }

        //========================================================================================

        public static List<Mate2> eListeContraintes(this Component2 cp, Boolean inclureSupprime = false)
        {
            List<Mate2> Liste = new List<Mate2>();

            Object[] TabContraintes = (Object[])cp.GetMates();
            if (TabContraintes.IsRef())
            {
                foreach (Mate2 Contrainte in TabContraintes)
                {
                    if (inclureSupprime)
                    {
                        Liste.Add(Contrainte);
                    }
                    else
                    {
                        Feature F = (Feature)Contrainte;
                        Boolean[] pArrayResult = (Boolean[])F.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, null);
                        if ((pArrayResult.IsRef()) && (Convert.ToBoolean(pArrayResult[0]) == false))
                            Liste.Add(Contrainte);
                    }
                }
            }

            return Liste;
        }

        public static List<Mate2> eChercherContraintes(this Component2 cp1, Component2 cp2, Boolean inclureSupprime = false)
        {
            List<Mate2> Liste = new List<Mate2>();

            foreach (Mate2 Contrainte in cp1.eListeContraintes(inclureSupprime))
            {
                foreach (MateEntity2 Ent in Contrainte.eListeDesEntitesDeContrainte())
                {
                    Component2 Cp = Ent.ReferenceComponent;

                    if (Cp.eNomSansExt() == cp2.eNomSansExt())
                        Liste.Add(Contrainte);
                }
            }

            return Liste;
        }

        public static List<MateEntity2> eListeDesEntitesDeContrainte(this Mate2 ct)
        {
            List<MateEntity2> Liste = new List<MateEntity2>();

            for (int i = 0; i < ct.GetMateEntityCount(); i++)
                Liste.Add(ct.MateEntity(i));

            return Liste;
        }

        //========================================================================================

        public static List<Body2> eListeCorps(this Component2 cp)
        {
            List<Body2> Liste = new List<Body2>();

            Object vInfosCorps;

            Object[] Corps = (Object[])cp.GetBodies3((int)swBodyType_e.swSolidBody, out vInfosCorps);
            int[] InfosCorps = (int[])vInfosCorps;

            for (int i = 0; i < Corps.Length; i++)
            {
                if (InfosCorps[i] == (int)swBodyInfo_e.swNormalBody_e)
                    Liste.Add((Body2)Corps[i]);
            }

            //foreach (Body2 Cp in Corps)
            //    Liste.Add(Cp);

            return Liste;
        }

        public static Body2 eChercherCorps(this Component2 cp, String nom, Boolean regex = true)
        {
            foreach (Body2 Corps in cp.eListeCorps())
            {
                if (regex)
                {
                    if (Regex.IsMatch(Corps.Name, nom))
                        return Corps;
                }
                else
                {
                    if (Corps.Name == nom)
                        return Corps;
                }
            }

            return null;
        }

        public static List<Body2> eListeCorps(this PartDoc piece, Boolean VisibleOnly)
        {
            List<Body2> Liste = new List<Body2>();

            Object[] Corps = (Object[])piece.GetBodies2((int)swBodyType_e.swSolidBody, VisibleOnly);

            foreach (Body2 Cp in Corps)
                Liste.Add(Cp);

            return Liste;
        }

        public static Body2 eChercherCorps(this PartDoc piece, String nomCorps, Boolean VisibleOnly)
        {
            Object[] Corps = (Object[])piece.GetBodies2((int)swBodyType_e.swSolidBody, VisibleOnly);

            foreach (Body2 c in Corps)
            {
                if (c.Name == nomCorps)
                    return c;
            }

            return null;
        }

        public static List<Body2> eListeDesCorps(this Feature f)
        {
            Dictionary<String, Body2> Dic = new Dictionary<String, Body2>();

            foreach (Face2 F in f.eListeDesFaces())
            {
                Body2 b = f.GetBody();
                if (b.IsNull()) continue;
                Dic.AddIfNotExist(b.Name, b);
            }

            return Dic.Values.ToList();
        }

        //========================================================================================

        public static List<Face2> eListeDesFaces(this Body2 c)
        {
            List<Face2> Liste = new List<Face2>();

            Object[] Faces = (Object[])c.GetFaces();

            foreach (Face2 F in Faces)
                Liste.Add(F);

            return Liste;
        }

        public static List<Face2> eListeDesFaces(this Edge e)
        {
            List<Face2> Liste = new List<Face2>();

            Object[] Faces = (Object[])e.GetTwoAdjacentFaces2();

            foreach (Face2 F in Faces)
                Liste.Add(F);

            return Liste;
        }

        public static List<Face2> eListeDesFaces(this Vertex v)
        {
            List<Face2> Liste = new List<Face2>();

            Object[] Faces = (Object[])v.GetAdjacentFaces();

            foreach (Face2 F in Faces)
                Liste.Add(F);

            return Liste;
        }

        public static List<Face2> eListeDesFacesModifiees(this Feature f)
        {
            List<Face2> Liste = new List<Face2>();

            Object[] Faces = (Object[])f.GetAffectedFaces();

            if (Faces != null)
            {
                foreach (Face2 F in Faces)
                    Liste.Add(F);
            }

            return Liste;
        }

        public static List<Face2> eListeDesFaces(this Feature f)
        {
            List<Face2> Liste = new List<Face2>();

            Object[] Faces = (Object[])f.GetFaces();

            if (Faces != null)
            {
                foreach (Face2 F in Faces)
                    Liste.Add(F);
            }

            return Liste;
        }

        public static List<Face2> eListeDesFacesContigues(this Face2 face)
        {
            List<Face2> ListeFaces = new List<Face2>();
            List<Edge> ListeArretes = face.eListeDesArretes();

            foreach (Edge E in ListeArretes)
            {
                foreach (Face2 F in E.eListeDesFaces())
                    ListeFaces.Add(F);

                ListeFaces.Remove(face);
            }

            return ListeFaces;
        }

        public static Boolean eFaceEstConnecte(this Face2 face, Face2 faceTest)
        {
            foreach (var f in face.eListeDesFacesContigues())
            {
                if (f.IsSame(faceTest))
                    return true;
            }

            return false;
        }

        public static Face2 eChercherFace(this Component2 cp, String nom)
        {
            ModelDoc2 Mdl = (ModelDoc2)cp.GetModelDoc2();

            foreach (Body2 C in cp.eListeCorps())
            {
                foreach (Face2 F in C.eListeDesFaces())
                {
                    String Nom = Mdl.GetEntityName(F);
                    if (Regex.IsMatch(Nom, nom))
                        return F;
                }
            }

            return null;
        }

        //========================================================================================

        public static List<Loop2> eListeDesBoucles(this Face2 f, Predicate<Loop2> filtreTest)
        {
            List<Loop2> Liste = new List<Loop2>();

            Object[] Loop = (Object[])f.GetLoops();

            foreach (Loop2 L in Loop)
            {
                if (filtreTest.IsNull() || filtreTest(L))
                {
                    Liste.Add(L);
                }
            }

            return Liste;
        }

        public static List<Loop2> eListeDesBoucles(this Face2 f)
        {
            return eListeDesBoucles(f, null);
        }

        public static List<CoEdge> eListeDesCoArrete(this Loop2 l)
        {
            List<CoEdge> Liste = new List<CoEdge>();

            Object[] CoEdge = (Object[])l.GetCoEdges();

            foreach (CoEdge C in CoEdge)
                Liste.Add(C);

            return Liste;
        }

        //========================================================================================

        public static Double eLgBoucle(this Loop2 l)
        {
            Double Lg = 0;
            foreach (var CoArrete in l.eListeDesCoArrete())
            {
                Edge Arrete = CoArrete.GetEdge();
                Lg += Arrete.eLgArrete();
            }

            return Lg;
        }

        public static Double eLgArrete(this Edge e)
        {
            Curve Courbe = e.GetCurve();
            double Start, End; bool Ferme, Periodic;
            Courbe.GetEndParams(out Start, out End, out Ferme, out Periodic);
            return Courbe.GetLength3(Start, End);
        }

        //========================================================================================

        public static Byte[] eGetPersistId(this Feature f, ModelDoc2 mdl)
        {
            return mdl.Extension.GetPersistReference3(f);
        }

        public static Boolean eIsSame(this Object obj1, Object obj2)
        {
            if (App.Sw.IsSame(obj1, obj2) == (int)swObjectEquality.swObjectSame)
                return true;

            return false;
        }

        public static Boolean eIsSamePID(this Object obj1, Object obj2, Component2 cp)
        {
            ModelDocExtension ext = cp.eModelDoc2().Extension;
            Object r1 = ext.GetPersistReference3(obj1);
            Object r2 = ext.GetPersistReference3(obj2);

            if (ext.IsSamePersistentID(r1, r2) == (int)swObjectEquality.swObjectSame)
                return true;

            return false;
        }

        public static Boolean eContient<T>(this List<T> liste, T Obj)
        {
            foreach (T o in liste)
            {
                swObjectEquality et = (swObjectEquality)App.Sw.IsSame(o, Obj);
                if (eIsSame(o, Obj))
                    return true;
            }

            return false;
        }

        public static List<T> eSupprimer<T>(this List<T> liste, Object Obj)
        {
            if (liste.Count == 0) return liste;

            int i = 0;
            T o = liste[i];
            while (o.IsRef())
            {
                if (o.eIsSame(Obj))
                    liste.RemoveAt(i);
                else
                    i++;

                if (i < liste.Count)
                    o = liste[i];
                else
                    break;
            }

            return liste;
        }

        //========================================================================================

        public static List<Edge> eListeDesArretes(this Face2 f)
        {
            List<Edge> Liste = new List<Edge>();

            Object[] Arretes = (Object[])f.GetEdges();

            foreach (Edge A in Arretes)
                Liste.Add(A);

            return Liste;
        }

        public static List<Edge> eListeDesArretes(this Edge e)
        {
            List<Edge> Liste = new List<Edge>();

            Vertex Vs = (Vertex)e.GetStartVertex();

            Object[] Arretes = (Object[])Vs.GetEdges();

            foreach (Edge A in Arretes)
                Liste.Add(A);

            Vs = (Vertex)e.GetEndVertex();

            Arretes = (Object[])Vs.GetEdges();

            foreach (Edge A in Arretes)
            {
                if (!Liste.Contains(A))
                    Liste.Add(A);
            }

            return Liste;
        }

        public static List<Edge> eListeDesArretesCommunes(this Face2 f1, Face2 f2)
        {
            List<Edge> Liste = new List<Edge>();
            List<Edge> ListeArrete = f1.eListeDesArretes();

            foreach (Edge E in f2.eListeDesArretes())
            {
                if (ListeArrete.eContient(E))
                    Liste.Add(E);
            }

            return Liste;
        }

        public static List<Edge> eListeDesArretesContigues(this Face2 face, Edge e1)
        {
            List<Edge> ListeArretes = face.eListeDesArretes();

            List<Edge> Liste = new List<Edge>();

            // On boucle sur les arretes de la face du dessus
            for (int i = 0; i < ListeArretes.Count; i++)
            {
                // Si on tombe sur l'arrete de base, on selectionne celle de devant
                // et celle de derrière.
                if (e1.eIsSame(ListeArretes[i]))
                {

                    // Si i-1 est inferieur à 0, on est au début de la liste
                    // donc on prend la dernière arete.
                    if ((i - 1) < 0)
                        Liste.Add(ListeArretes[ListeArretes.Count - 1]);
                    else
                        Liste.Add(ListeArretes[i - 1]);

                    // Si i+1 est superieur à Count, on est à la fin de la liste
                    // donc on prend la première arete.
                    if ((i + 1) < ListeArretes.Count)
                        Liste.Add(ListeArretes[i + 1]);
                    else
                        Liste.Add(ListeArretes[0]);

                    break;
                }
            }

            return Liste;
        }

        //========================================================================================

        public static Boolean eSelectEntite(this Entity entite, Boolean ajouter = false)
        {
            return entite.Select4(ajouter, null);
        }

        public static Boolean eSelectEntite(this Entity entite, ModelDoc2 mdl, int marque = -1, Boolean ajouter = false)
        {
            SelectionMgr SelMgr = mdl.SelectionManager;
            SelectData SelData = SelMgr.CreateSelectData();
            SelData.Mark = marque;
            return entite.Select4(ajouter, SelData);
        }

        public static Boolean eSelectEntite(this Face2 face, Boolean ajouter = false)
        {
            return eSelectEntite((Entity)face, ajouter);
        }

        public static Boolean eSelectEntite(this Face2 face, ModelDoc2 mdl, int marque = -1, Boolean ajouter = false)
        {
            return eSelectEntite((Entity)face, mdl, marque, ajouter);
        }

        public static Boolean eSelectEntite(this Edge bord, Boolean ajouter = false)
        {
            return eSelectEntite((Entity)bord, ajouter);
        }

        public static Boolean eSelectEntite(this Edge bord, ModelDoc2 mdl, int marque = -1, Boolean ajouter = false)
        {
            return eSelectEntite((Entity)bord, mdl, marque, ajouter);
        }

        public static Boolean eSelectEntite(this Vertex point, Boolean ajouter = false)
        {
            return eSelectEntite((Entity)point, ajouter);
        }

        public static Boolean eSelectEntite(this Vertex point, ModelDoc2 mdl, int marque = -1, Boolean ajouter = false)
        {
            return eSelectEntite((Entity)point, mdl, marque, ajouter);
        }

        /// <summary>
        /// Selectionner une fonction.
        /// Ne fonctionne pas pendant l'edition d'un PropertyManagerPage
        /// </summary>
        /// <param name="f"></param>
        /// <param name="ajouter"></param>
        public static Boolean eSelect(this Feature f, Boolean ajouter = false)
        {
            return f.Select2(ajouter, -1);
        }

        /// <summary>
        /// Selectionner une fonction.
        /// Ne fonctionne pas pendant l'edition d'un PropertyManagerPage
        /// </summary>
        /// <param name="f"></param>
        /// <param name="marque"></param>
        /// <param name="ajouter"></param>
        public static Boolean eSelect(this Feature f, int marque, Boolean ajouter = false)
        {
            return f.Select2(ajouter, marque);
        }

        public static Boolean eSelect(this Body2 c, Boolean ajouter = false)
        {
            return c.Select2(ajouter, null);
        }

        public static Boolean eSelect(this Body2 c, ModelDoc2 mdl, int marque, Boolean ajouter = false)
        {
            SelectionMgr SelMgr = mdl.SelectionManager;
            SelectData SelData = SelMgr.CreateSelectData();
            SelData.Mark = marque;
            return c.Select2(ajouter, SelData);
        }

        /// <summary>
        /// Selectionner un composant.
        /// Ne fonctionne pas pendant l'edition d'un PropertyManagerPage
        /// </summary>
        /// <param name="c"></param>
        /// <param name="mdl"></param>
        /// <param name="marque"></param>
        /// <param name="ajouter"></param>
        public static Boolean eSelect(this Component2 c, Boolean ajouter = false)
        {
            return c.Select4(ajouter, null, false);
        }

        /// <summary>
        /// Selectionner un composant.
        /// Ne fonctionne pas pendant l'edition d'un PropertyManagerPage
        /// </summary>
        /// <param name="c"></param>
        /// <param name="mdl"></param>
        /// <param name="marque"></param>
        /// <param name="ajouter"></param>
        public static Boolean eSelect(this Component2 c, ModelDoc2 mdl, int marque, Boolean ajouter = false)
        {
            SelectionMgr SelMgr = mdl.SelectionManager;
            SelectData SelData = SelMgr.CreateSelectData();
            SelData.Mark = marque;
            return c.Select4(ajouter, SelData, false);
        }

        /// <summary>
        /// Extension de SelectById2
        /// </summary>
        /// <param name="c"></param>
        /// <param name="mdl"></param>
        /// <param name="marque"></param>
        /// <param name="ajouter"></param>
        public static Boolean eSelectById(this Component2 c, ModelDoc2 mdl, int marque = -1, Boolean ajouter = false, Boolean forceSelection = false)
        {
            Boolean t = mdl.Extension.SelectByID2(c.GetSelectByIDString(), swSelectType.swSelCOMPONENTS, 0, 0, 0, ajouter, marque, null, (int)swSelectOption_e.swSelectOptionDefault);
            if (!t && forceSelection)
                t = mdl.Extension.SelectByID2(c.GetSelectByIDString(), swSelectType.swSelCOMPONENTS, 0, 0, 0, ajouter, marque, null, (int)swSelectOption_e.swSelectOptionDefault);

            return t;
        }

        /// <summary>
        /// Extension de DeSelectByID
        /// </summary>
        /// <param name="c"></param>
        /// <param name="mdl"></param>
        public static Boolean eDeSelectById(this Component2 c, ModelDoc2 mdl)
        {
            return mdl.DeSelectByID(c.GetSelectByIDString(), swSelectType.swSelCOMPONENTS, 0, 0, 0);
        }

        /// <summary>
        /// Extension de SelectById2
        /// </summary>
        /// <param name="mdl"></param>
        /// <param name="NomComposant"></param>
        /// <param name="marque"></param>
        /// <param name="ajouter"></param>
        /// <returns></returns>
        public static Boolean eSelectByIdComp(this ModelDoc2 mdl, String NomComposant, int marque = -1, Boolean ajouter = false)
        {
            return mdl.Extension.SelectByID2(NomComposant, swSelectType.swSelCOMPONENTS, 0, 0, 0, ajouter, marque, null, (int)swSelectOption_e.swSelectOptionDefault);
        }

        public static int eSelectMulti<T>(this ModelDoc2 mdl, List<T> lst, int? marque, Boolean ajouter = false)
            where T : class
        {
            SelectionMgr SelMgr = mdl.SelectionManager;
            SelectData SelectData = default(SelectData);
            if (marque != null)
            {
                SelectData = SelMgr.CreateSelectData();
                SelectData.Mark = (int)marque;
            }
            else
                SelectData = null;

            return mdl.Extension.MultiSelect2(lst.ToArray(), ajouter, SelectData);
        }

        public static int eSelectMulti<T>(this ModelDoc2 mdl, T obj, int? marque, Boolean ajouter = false)
            where T : class
        {
            return mdl.eSelectMulti(new List<T>() { obj }, marque, ajouter);
        }

        public static void eEffacerSelection(this ModelDoc2 mdl)
        {
            mdl.ClearSelection2(true);
        }

        /// <summary>
        /// Selectionner une fonction.
        /// Fonctionne dans tout les cas
        /// </summary>
        /// <param name="f"></param>
        /// <param name="mdl"></param>
        /// <param name="marque"></param>
        /// <param name="ajouter"></param>
        public static Boolean eSelectionnerPMP(this Feature f, ModelDoc2 mdl, int marque = -1, Boolean ajouter = false)
        {
            String TypeF;
            String Nom = f.GetNameForSelection(out TypeF);
            return mdl.Extension.SelectByID2(Nom, TypeF, 0, 0, 0, ajouter, marque, null, (int)swSelectOption_e.swSelectOptionDefault);
        }

        /// <summary>
        /// Index de base 1
        /// </summary>
        /// <param name="mdl"></param>
        /// <param name="index"></param>
        /// <param name="marque"></param>
        /// <returns></returns>
        public static swSelectType_e eSelect_RecupererTypeObjet(this ModelDoc2 mdl, int index = 1, int marque = -1)
        {
            SelectionMgr SelMgr = mdl.SelectionManager;
            return (swSelectType_e)SelMgr.GetSelectedObjectType3(index, marque);
        }

        /// <summary>
        /// Index de base 1
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="mdl"></param>
        /// <param name="index"></param>
        /// <param name="marque"></param>
        /// <returns></returns>
        public static T eSelect_RecupererObjet<T>(this ModelDoc2 mdl, int index = 1, int marque = -1)
            where T : class
        {
            SelectionMgr SelMgr = mdl.SelectionManager;

            if (SelMgr.GetSelectedObjectCount2(marque) == 0)
                return null;

            return SelMgr.GetSelectedObject6(index, marque) as T;
        }

        public static List<T> eSelect_RecupererListeObjets<T>(this ModelDoc2 mdl, int marque = -1)
            where T : class
        {
            List<T> Liste = new List<T>();

            SelectionMgr SelMgr = mdl.SelectionManager;
            int Count = SelMgr.GetSelectedObjectCount2(marque);

            for (int i = 1; i <= Count; i++)
                Liste.Add(SelMgr.GetSelectedObject6(i, marque) as T);

            return Liste;
        }

        public static List<Component2> eSelect_RecupererListeComposants(this ModelDoc2 mdl, int marque = -1)
        {
            List<Component2> Liste = new List<Component2>();

            SelectionMgr SelMgr = mdl.SelectionManager;
            int Count = SelMgr.GetSelectedObjectCount2(marque);

            for (int i = 1; i <= Count; i++)
                Liste.Add(SelMgr.GetSelectedObjectsComponent4(i, marque) as Component2);

            return Liste;
        }

        /// <summary>
        /// Index de base 1
        /// </summary>
        /// <param name="mdl"></param>
        /// <param name="index"></param>
        /// <param name="marque"></param>
        /// <returns></returns>
        public static Component2 eSelect_RecupererComposant(this ModelDoc2 mdl, int index = 1, int marque = -1)
        {
            SelectionMgr SelMgr = mdl.SelectionManager;
            if (SelMgr.GetSelectedObjectCount2(marque) == 0)
                return null;

            return SelMgr.GetSelectedObjectsComponent4(index, marque);
        }

        //========================================================================================

        public static string eGetNomEntite(this Face2 ent, ModelDoc2 modele)
        {
            return modele.GetEntityName(ent);
        }

        public static string eGetNomEntite(this Edge ent, ModelDoc2 modele)
        {
            return modele.GetEntityName(ent);
        }

        public static string eGetNomEntite(this Vertex ent, ModelDoc2 modele)
        {
            return modele.GetEntityName(ent);
        }


        private static string eSetNomEntite(Object o, String nom, ModelDoc2 modele)
        {
            PartDoc Doc = modele as PartDoc;
            if (Doc == null)
                return "";

            Doc.SetEntityName(o, nom);

            return Doc.GetEntityName(o);
        }

        public static string eSetNomEntite(this Face2 ent, String nom, ModelDoc2 modele)
        {
            return eSetNomEntite((Object)ent, nom, modele);
        }

        public static string eSetNomEntite(this Edge ent, String nom, ModelDoc2 modele)
        {
            return eSetNomEntite((Object)ent, nom, modele);
        }

        public static string eSetNomEntite(this Vertex ent, String nom, ModelDoc2 modele)
        {
            return eSetNomEntite((Object)ent, nom, modele);
        }

        //========================================================================================

        public static Double eVolume(this Body2 corps)
        {
            Double[] pProps = (Double[])corps.GetMassProperties(1);
            Double pVolume = pProps[3];

            return pVolume;
        }

        /// <summary>
        /// Renvoi le point extreme dans la direction donnée
        /// </summary>
        /// <param name="corps"></param>
        /// <param name="direction"></param>
        /// <returns></returns>
        public static Point ePointExtreme(this Body2 corps, Vecteur direction)
        {
            Double oX = 0, oY = 0, oZ = 0;
            corps.GetExtremePoint(direction.X, direction.Y, direction.Z, out oX, out oY, out oZ);
            return new Point(oX, oY, oZ);
        }

        /// <summary>
        /// Centre de gravité du Corps
        /// </summary>
        /// <param name="corps"></param>
        /// <returns></returns>
        public static Double[] eCdG(this Body2 corps)
        {
            Double[] pProps = (Double[])corps.GetMassProperties(1);
            Double[] pCdG = new Double[] { pProps[0], pProps[1], pProps[2] };

            return pCdG;
        }

        /// <summary>
        /// Verification de la similitude de deux corps et exclu symetrie
        /// </summary>
        /// <param name="corps"></param>
        /// <param name="corpsTest"></param>
        /// <returns></returns>
        public static Boolean eEstSemblable(this Body2 corps, Body2 corpsTest)
        {
            MathTransform mt = null;

            return corps.eEstSemblable(corpsTest, out mt);
        }

        /// <summary>
        /// Verification de la similitude de deux corps et exclu symetrie
        /// </summary>
        /// <param name="corps"></param>
        /// <param name="corpsTest"></param>
        /// <param name="mt"></param>
        /// <returns></returns>
        public static Boolean eEstSemblable(this Body2 corps, Body2 corpsTest, out MathTransform mt)
        {
            // La méthode renvoi également true si les corps sont symétriques.
            Boolean result = corps.GetCoincidenceTransform2((Object)corpsTest, out mt);

            if (result == true)
            {
                // Vérification du déterminant de la matrice de rotation
                // Calcul du déterminant
                double[] v = (double[])mt.ArrayData;
                double det1 = v[0] * ((v[4] * v[8]) - (v[7] * v[5]));
                double det2 = v[1] * ((v[3] * v[8]) - (v[6] * v[5]));
                double det3 = v[2] * ((v[3] * v[7]) - (v[6] * v[4]));
                double det = det1 - det2 + det3;

                // Si le déterminant est == -1, la matrice est une symetrie
                // Les corps ne sont pas semblables
                if (det < 0)
                    result = false;
            }

            return result;
        }

        public static Boolean eVisible(this Body2 corps)
        {
            return !corps.DisableDisplay;
        }

        public static void eVisible(this Body2 corps, Boolean visible)
        {
            corps.DisableDisplay = !visible;
            corps.HideBody(!visible);
        }

        public static void eVisible(this Body2 corps, PartDoc piece, Boolean visible)
        {
            corps.DisableDisplay = !visible;
            corps.HideBody(!visible);

            corps.Select2(false, null);

            if (visible)
                piece.eModelDoc2().FeatureManager.ShowBodies();
            else
                piece.eModelDoc2().FeatureManager.HideBodies();

            piece.eModelDoc2().ClearSelection2(true);
        }

        public static eTypeCorps eTypeDeCorps(this Body2 corps)
        {
            if (corps.IsRef())
            {
                if (corps.IsSheetMetal())
                    return eTypeCorps.Tole;

                foreach (Feature Fonction in corps.GetFeatures())
                {
                    switch (Fonction.GetTypeName2())
                    {
                        case "SheetMetal":
                        case "SMBaseFlange":
                        case "SolidToSheetMetal":
                        case "FlatPattern":
                            return eTypeCorps.Tole;
                        case "WeldMemberFeat":
                        case "WeldCornerFeat":
                            return eTypeCorps.Barre;
                        default:
                            break;
                    }
                }
            }
            return eTypeCorps.Autre;
        }

        public static Boolean eSetMateriau(this Body2 corps, String nomMateriau, String nomConfig)
        {
            String BaseMat = eBaseDuMateriau(nomMateriau);

            int result = corps.SetMaterialProperty(nomConfig, BaseMat, nomMateriau);

            if (result != (int)swBodyMaterialApplicationError_e.swBodyMaterialApplicationError_NoError)
                return false;

            return true;
        }

        public static Boolean eSetMateriau(this Body2 corps, String nomMateriau)
        {
            String nomConfig = eModeleActif().eNomConfigActive();
            String BaseMat = eBaseDuMateriau(nomMateriau);

            int result = corps.SetMaterialProperty(nomConfig, BaseMat, nomMateriau);

            if (result != (int)swBodyMaterialApplicationError_e.swBodyMaterialApplicationError_NoError)
                return false;

            return true;
        }

        public static String eGetMateriau(this Body2 corps, String config, out String baseMateriau)
        {
            String Materiau = corps.GetMaterialPropertyName(config, out baseMateriau);
            if (String.IsNullOrEmpty(Materiau))
                Materiau = "";

            return Materiau;
        }

        public static String eGetMateriau(this Body2 corps, String config)
        {
            String Db = "";
            String Materiau = corps.GetMaterialPropertyName(config, out Db);
            if (String.IsNullOrEmpty(Materiau))
                Materiau = "";

            return Materiau;
        }

        public static String eGetMateriau(this Body2 corps, PartDoc piece, out String baseMateriau)
        {

            String pNomConfigActive = piece.eModelDoc2().eNomConfigActive();
            String Materiau = corps.GetMaterialPropertyName(pNomConfigActive, out baseMateriau);
            if (String.IsNullOrEmpty(Materiau))
                Materiau = piece.eGetMateriau();

            return Materiau;
        }

        public static String eGetMateriau(this Body2 corps, PartDoc piece)
        {
            String Db = "";
            String pNomConfigActive = piece.eModelDoc2().eNomConfigActive();
            String Materiau = corps.GetMaterialPropertyName(pNomConfigActive, out Db);
            if (String.IsNullOrEmpty(Materiau))
                Materiau = piece.eGetMateriau();

            return Materiau;
        }

        public static String eGetMateriauCorpsOuComp(this Body2 corps, Component2 cp)
        {
            String Db = "";
            String Materiau = corps.eGetMateriau(cp.eNomConfiguration(), out Db);
            if (String.IsNullOrWhiteSpace(Materiau))
                Materiau = cp.eGetMateriau(out Db);

            return Materiau;
        }

        public static String eGetMateriauCorpsOuComp(this Body2 corps, Component2 cp, out String baseMateriau)
        {
            String Materiau = corps.eGetMateriau(cp.eNomConfiguration(), out baseMateriau);
            if (String.IsNullOrWhiteSpace(Materiau))
                Materiau = cp.eGetMateriau(out baseMateriau);

            return Materiau;
        }

        public static String eGetMateriauCorpsOuPiece(this Body2 corps, PartDoc piece, String nomConfig)
        {
            String Db = "";
            String Materiau = corps.eGetMateriau(nomConfig, out Db);
            if (String.IsNullOrWhiteSpace(Materiau))
                Materiau = piece.eGetMateriau(nomConfig, out Db);

            return Materiau;
        }

        public static String eGetMateriauCorpsOuPiece(this Body2 corps, PartDoc piece, String nomConfig, out String baseMateriau)
        {
            String Materiau = corps.eGetMateriau(nomConfig, out baseMateriau);
            if (String.IsNullOrWhiteSpace(Materiau))
                Materiau = piece.eGetMateriau(nomConfig, out baseMateriau);

            return Materiau;
        }

        public static String eGetMateriau(this PartDoc piece, String config, out String baseMateriau)
        {
            String pMateriau = piece.GetMaterialPropertyName2(config, out baseMateriau);
            if (String.IsNullOrEmpty(pMateriau))
                pMateriau = "";
            return pMateriau;
        }

        public static String eGetMateriau(this PartDoc piece, out String baseMateriau)
        {
            String pMateriau = piece.GetMaterialPropertyName2(piece.eModelDoc2().eNomConfigActive(), out baseMateriau);
            if (String.IsNullOrEmpty(pMateriau))
                pMateriau = "";
            return pMateriau;
        }

        public static String eGetMateriau(this PartDoc piece, String config)
        {
            String Db = "";
            String pMateriau = piece.GetMaterialPropertyName2(config, out Db);
            if (String.IsNullOrEmpty(pMateriau))
                pMateriau = "";
            return pMateriau;
        }

        public static String eGetMateriau(this PartDoc piece)
        {
            String Db = "";
            String pMateriau = piece.GetMaterialPropertyName2(piece.eModelDoc2().eNomConfigActive(), out Db);
            if (String.IsNullOrEmpty(pMateriau))
                pMateriau = "";
            return pMateriau;
        }

        public static String eGetMateriau(this Component2 cp, out String baseMateriau)
        {
            String pMateriau = cp.ePartDoc().GetMaterialPropertyName2(cp.eNomConfiguration(), out baseMateriau);
            if (String.IsNullOrEmpty(pMateriau))
                pMateriau = "";
            return pMateriau;
        }

        public static String eGetMateriau(this Component2 cp)
        {
            String Db = "";
            String pMateriau = cp.ePartDoc().GetMaterialPropertyName2(cp.eNomConfiguration(), out Db);
            if (String.IsNullOrEmpty(pMateriau))
                pMateriau = "";
            return pMateriau;
        }

        public static String eGetMateriau(this BodyFolder dossier)
        {
            String materiau = dossier.eProp("Materiau");

            if (String.IsNullOrWhiteSpace(materiau))
                materiau = dossier.eProp("Matériau");

            return materiau;
        }

        public static void eSetMateriau(this PartDoc piece, String nomMateriau, String baseMateriau, String nomConfig)
        {
            piece.SetMaterialPropertyName2(nomConfig, baseMateriau, nomMateriau);
        }

        public static void eSetMateriau(this PartDoc piece, String nomMateriau, String baseMateriau)
        {
            piece.eSetMateriau(nomMateriau, baseMateriau, piece.eModelDoc2().eNomConfigActive());
        }

        public static void eSetMateriau(this PartDoc piece, String nomMateriau)
        {
            piece.eSetMateriau(nomMateriau, eBaseDuMateriau(nomMateriau), piece.eModelDoc2().eNomConfigActive());
        }

        /// <summary>
        /// Liste des bases de materiaux dans SW
        /// </summary>
        /// <returns></returns>
        public static List<String> eListeDesBasesMateriaux()
        {
            String[] Tab = App.Sw.GetMaterialDatabases();

            return Tab.ToList();
        }

        public static List<String> eListeDesBasesMateriaux(String nomBase)
        {
            String[] Tab = App.Sw.GetMaterialDatabases();

            List<String> Liste = new List<String>();

            foreach (String c in Tab)
            {
                if (Path.GetFileNameWithoutExtension(c).ToLowerInvariant() == nomBase.ToLowerInvariant())
                    Liste.Add(c);
            }

            return Liste;
        }

        /// <summary>
        /// Liste des base de materiaux contenant le materiau spécifié
        /// </summary>
        /// <param name="materiau"></param>
        /// <returns></returns>
        public static List<String> eListeBaseDuMateriau(String materiau)
        {
            List<String> ListeBaseMat = new List<string>();

            foreach (String Base in eListeDesBasesMateriaux())
            {
                using (XmlReader reader = XmlReader.Create(Base))
                {
                    while (reader.ReadToFollowing("material"))
                    {
                        // first element is the root element
                        if (reader.NodeType == XmlNodeType.Element)
                        {
                            reader.MoveToAttribute("name");
                            String Mat = reader.Value;
                            if (materiau == Mat)
                            {
                                ListeBaseMat.Add(Base);
                                break;
                            }
                        }
                    }
                }

            }

            return ListeBaseMat;
        }

        /// <summary>
        /// Renvoi la première base contenant le materiau spécifié
        /// </summary>
        /// <param name="materiau"></param>
        /// <returns></returns>
        public static String eBaseDuMateriau(String materiau)
        {
            List<string> Liste = eListeBaseDuMateriau(materiau);
            if (Liste.Count > 0)
                return Liste[0];

            return "";
        }

        public static Double eProprieteMat(String nomBase, String materiau, eMatPropriete prop)
        {
            String C;

            return eProprieteMat(nomBase, materiau, prop, out C);
        }

        public static Double eProprieteMat(String nomBase, String materiau, eMatPropriete prop, out String classe)
        {
            try
            {
                String CheminBase = eListeDesBasesMateriaux(nomBase)[0];

                using (XmlReader ClasseReader = XmlReader.Create(CheminBase))
                {
                    while (ClasseReader.ReadToFollowing("classification"))
                    {
                        using (XmlReader MatReader = ClasseReader.ReadSubtree())
                        {
                            ClasseReader.MoveToAttribute("name");
                            classe = ClasseReader.Value;

                            while (MatReader.ReadToFollowing("material"))
                            {
                                MatReader.MoveToAttribute("name");
                                if (MatReader.Value == materiau)
                                {
                                    MatReader.ReadToFollowing(prop.GetEnumInfo<TagXml>());

                                    MatReader.MoveToAttribute("value");
                                    return Convert.ToDouble(MatReader.Value);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            { Log.Message(e); }

            classe = "";
            return 0.0;
        }

        //========================================================================================

        public static Feature eFonctionTolerie(this Body2 corps)
        {
            foreach (Feature pFonc in corps.GetFeatures())
            {
                if (pFonc.GetTypeName2() == FeatureType.swTnSheetMetal)
                    return pFonc;
            }

            return null;
        }

        public static Feature eFonctionEtatDepliee(this Body2 corps)
        {
            foreach (Feature pFonc in corps.GetFeatures())
            {
                if (pFonc.GetTypeName2() == FeatureType.swTnFlatPattern)
                    return pFonc;
            }

            return null;
        }

        public static Body2 eCorpsDeTolerie(this BodyFolder dossier)
        {
            foreach (Body2 c in dossier.GetBodies())
            {
                if (c.eTypeDeCorps() == eTypeCorps.Tole)
                    return c;
            }

            return null;
        }

        public static String eNomConfigDepliee(String nomConfigPliee, String refDossier)
        {
            return nomConfigPliee + CONSTANTES.CONFIG_DEPLIEE + "_" + refDossier;
        }

        public static String eNomConfigDepliee(String nomConfigPliee, BodyFolder dossier)
        {
            String no = dossier.eProp(CONSTANTES.REF_DOSSIER);
            if (String.IsNullOrWhiteSpace(no))
                return null;

            return eNomConfigDepliee(nomConfigPliee, no);
        }

        /// <summary>
        /// Epaisseur en mm
        /// Issue de la fonction tolerie du 1er corps ou le cas echéant du dossier
        /// </summary>
        /// <param name="dossier"></param>
        /// <returns></returns>
        public static Double eEpaisseur1ErCorpsOuDossier(this BodyFolder dossier)
        {
            Body2 corps = dossier.ePremierCorps();
            Double E = corps.eEpaisseurCorps();
            if (E == -1)
                E = dossier.eEpaisseurDossier();

            return E;
        }

        /// <summary>
        /// Epaisseur en mm
        /// Issue de la fonction tolerie ou le cas echéant du dossier
        /// </summary>
        /// <param name="corps"></param>
        /// <param name="dossier"></param>
        /// <returns></returns>
        public static Double eEpaisseurCorpsOuDossier(this Body2 corps, BodyFolder dossier)
        {
            Double E = corps.eEpaisseurCorps();
            if (E == -1)
                E = dossier.eEpaisseurDossier();

            return E;
        }

        /// <summary>
        /// Epaisseur en mm
        /// Issue de la fonction tolerie
        /// </summary>
        /// <param name="corps"></param>
        /// <returns></returns>
        public static Double eEpaisseurCorps(this Body2 corps)
        {
            Double Ep = -1;
            try
            {
                SheetMetalFeatureData pParam = corps.eFonctionTolerie().GetDefinition();
                Ep = Math.Round(pParam.Thickness * 1000, 5);
            }
            catch (Exception e)
            {
                Log.Message(e);
                Log.MessageF("Corps : {0}", corps.Name);
            }

            return Ep;
        }

        /// <summary>
        /// Epaisseur en mm
        /// Issue du dossier
        /// </summary>
        /// <param name="dossier"></param>
        /// <returns></returns>
        public static Double eEpaisseurDossier(this BodyFolder dossier)
        {
            Double Ep = -1;
            try
            {
                if (dossier.ePropExiste(CONSTANTES.EPAISSEUR_DE_TOLERIE))
                    Ep = dossier.eProp(CONSTANTES.EPAISSEUR_DE_TOLERIE).eToDouble();
            }
            catch (Exception e)
            { Log.Message(e); }

            return Ep;
        }

        /// <summary>
        /// Rayon interieur des plis en mm
        /// </summary>
        public static Double eRayon(this Body2 corps)
        {
            SheetMetalFeatureData pParam = corps.eFonctionTolerie().GetDefinition();
            return Math.Round(pParam.BendRadius * 1000, 5);
        }

        /// <summary>
        /// Facteur K
        /// </summary>
        public static Double eFacteurK(this Body2 corps)
        {
            SheetMetalFeatureData pParam = corps.eFonctionTolerie().GetDefinition();
            return Math.Round(pParam.KFactor, 5);
        }

        public static ModelDoc2 eEnregistrerSous(this Body2 corps, PartDoc piece, String dossier, String nomFichier, eTypeFichierExport typeExport)
        {
            int pStatut;
            int pWarning;
            corps.Select2(false, null);

            Boolean Resultat = piece.SaveToFile3(Path.Combine(dossier, nomFichier + typeExport.GetEnumInfo<ExtFichier>()),
                                                  (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                                  (int)swCutListTransferOptions_e.swCutListTransferOptions_CutListProperties,
                                                  false,
                                                  "",
                                                  out pStatut,
                                                  out pWarning);
            if (Resultat)
                return App.ModelDoc2;

            return null;
        }

        public static String eProfilDossier(this BodyFolder dossier)
        {
            String Profil = "";
            try
            {
                if (dossier.ePropExiste(CONSTANTES.PROFIL_NOM))
                    Profil = dossier.eProp(CONSTANTES.PROFIL_NOM);
            }
            catch (Exception e)
            { Log.Message(e); }

            return Profil;
        }

        public static int eNbIntersection(this Component2 compBase, Body2 corpsBase, Component2 compTest, Body2 corpsTest)
        {
            int NbInt = 0;

            MathTransform XFormBase = compBase.Transform2;
            MathTransform XFormTest = compTest.Transform2;

            Body2 copieCorpsBase = corpsBase.Copy();
            Body2 copieCorpsTest = corpsTest.Copy();

            copieCorpsBase.ApplyTransform(XFormBase);
            copieCorpsTest.ApplyTransform(XFormTest);

            int Err;
            Object[] ListeCorpsIntersection = copieCorpsBase.Operations2((int)swBodyOperationType_e.SWBODYINTERSECT, copieCorpsTest, out Err);

            if (ListeCorpsIntersection != null)
                NbInt = ListeCorpsIntersection.Length;

            ListeCorpsIntersection = null;

            copieCorpsBase = null;
            copieCorpsTest = null;
            XFormBase = null;
            XFormTest = null;

            return NbInt;
        }

        //========================================================================================

        public static List<View> eListeDesVues(this Sheet feuille)
        {
            List<View> pListeVues = new List<View>();

            Object[] pTabVues = feuille.GetViews();

            if (pTabVues == null)
                return pListeVues;

            foreach (View pSwVue in pTabVues)
            {
                pListeVues.Add(pSwVue);
            }

            return pListeVues;
        }

        public static void eParcourirLesVues(this Sheet feuille, Predicate<View> filtre)
        {
            Object[] pTabVues = feuille.GetViews();

            if (pTabVues == null)
                return;

            foreach (View pSwVue in pTabVues)
            {
                if (filtre(pSwVue))
                    break;
            }
        }

        public static eZone eZoneVue(this View vue)
        {
            eZone e = new eZone();

            Double[] pArr = vue.GetOutline();

            e.PointMin.X = pArr[0];
            e.PointMin.Y = pArr[1];
            e.PointMax.X = pArr[2];
            e.PointMax.Y = pArr[3];

            return e;
        }

        public static void eSelectionner(this View vue, DrawingDoc dessin, Boolean ajouter = false)
        {
            ModelDoc2 mdl = dessin.eModelDoc2();
            ePoint p = vue.eZoneVue().CentreZone;
            mdl.Extension.SelectByID2(vue.Name, "DRAWINGVIEW", p.X, p.Y, p.Z, ajouter, -1, null, 0);
        }

        public static eZone eEnveloppeDesVues(this Sheet feuille)
        {
            eZone e = new eZone();

            e.PointMax.X = Double.NegativeInfinity;
            e.PointMax.Y = Double.NegativeInfinity;
            e.PointMin.X = Double.PositiveInfinity;
            e.PointMin.Y = Double.PositiveInfinity;


            feuille.eParcourirLesVues(
                    v =>
                    {
                        eZone d = v.eZoneVue();

                        e.PointMax.X = Math.Max(e.PointMax.X, d.PointMax.X);
                        e.PointMax.Y = Math.Max(e.PointMax.Y, d.PointMax.Y);
                        e.PointMin.X = Math.Min(e.PointMin.X, Math.Max(Double.NegativeInfinity, d.PointMin.X));
                        e.PointMin.Y = Math.Min(e.PointMin.Y, Math.Max(Double.NegativeInfinity, d.PointMin.Y));

                        return false;
                    }
                    );

            return e;
        }

        public static eOrientation eGetOrientation(this Sheet feuille)
        {

            Double pLargeur = 0;
            Double pHauteur = 0;
            swDwgPaperSizes_e pTaille = (swDwgPaperSizes_e)feuille.GetSize(ref pLargeur, ref pHauteur);
            if (pLargeur > pHauteur)
                return eOrientation.Paysage;

            return eOrientation.Portrait;
        }

        public static void eSetOrientation(this Sheet feuille, eOrientation orientation)
        {
            Double pLargeur = 0;
            Double pHauteur = 0;
            swDwgPaperSizes_e pTaille = (swDwgPaperSizes_e)feuille.GetSize(ref pLargeur, ref pHauteur);

            if (feuille.eGetOrientation() != orientation)
                feuille.SetSize(12, pHauteur, pLargeur);
        }

        public static eFormat eGetFormat(this Sheet feuille)
        {
            Double pLargeur = 0;
            Double pHauteur = 0;
            swDwgPaperSizes_e pTaille = (swDwgPaperSizes_e)feuille.GetSize(ref pLargeur, ref pHauteur);

            if (feuille.eGetOrientation() == eOrientation.Portrait)
            {
                Double pTmp = pLargeur;
                pLargeur = pHauteur;
                pHauteur = pTmp;
            }

            if ((pLargeur == 1.189) && (pHauteur == 0.841))
                return eFormat.A0;

            if ((pLargeur == 0.841) && (pHauteur == 0.594))
                return eFormat.A1;

            if ((pLargeur == 0.594) && (pHauteur == 0.420))
                return eFormat.A2;

            if ((pLargeur == 0.420) && (pHauteur == 0.297))
                return eFormat.A3;

            if ((pLargeur == 0.297) && (pHauteur == 0.210))
                return eFormat.A4;

            if ((pLargeur == 0.210) && (pHauteur == 0.148))
                return eFormat.A5;

            return eFormat.Utilisateur;
        }

        public static void eSetFormat(this Sheet feuille, eFormat format)
        {

            Double pLargeur = 0;
            Double pHauteur = 0;

            switch (format)
            {
                case eFormat.A0:
                    pLargeur = 1.189; pHauteur = 0.841;
                    break;
                case eFormat.A1:
                    pLargeur = 0.841; pHauteur = 0.594;
                    break;
                case eFormat.A2:
                    pLargeur = 0.594; pHauteur = 0.420;
                    break;
                case eFormat.A3:
                    pLargeur = 0.420; pHauteur = 0.297;
                    break;
                case eFormat.A4:
                    pLargeur = 0.297; pHauteur = 0.210;
                    break;
                case eFormat.A5:
                    pLargeur = 0.210; pHauteur = 0.148;
                    break;
            }

            if (feuille.eGetOrientation() == eOrientation.Portrait)
            {
                Double pTmp = pLargeur;
                pLargeur = pHauteur;
                pHauteur = pTmp;
            }

            feuille.SetSize(12, pLargeur, pHauteur);
        }

        public static String eGetGabaritDeFeuille(this Sheet feuille)
        {
            return feuille.GetTemplateName();
        }

        public static void eSetGabaritDeFeuille(this Sheet feuille, String gabarit)
        {
            feuille.SetTemplateName(gabarit);
        }

        public static void eAjusterAutourDesVues(this Sheet feuille)
        {
            eZone pEnveloppe = feuille.eEnveloppeDesVues();

            if (pEnveloppe == null)
                return;

            feuille.eRedimensionner(pEnveloppe.PointMax.X + Math.Max(0, pEnveloppe.PointMin.X), pEnveloppe.PointMax.Y + Math.Max(0, pEnveloppe.PointMin.Y));
        }

        public static void eRedimensionner(this Sheet feuille, Double largeur, Double hauteur)
        {

            if ((largeur == 0) || (hauteur == 0))
                return;

            feuille.SetSize((int)swDwgPaperSizes_e.swDwgPapersUserDefined, largeur, hauteur);

        }

        public static String eExporterEn(this Sheet feuille, DrawingDoc dessin, eTypeFichierExport TypeExport, String CheminDossier, String NomDuFichierAlternatif = "", Boolean ToutesLesFeuilles = false)
        {
            ExportPdfData OptionsPDF = null;

            switch (TypeExport)
            {
                case eTypeFichierExport.DXF:
                    Sw.DxfDwg_ExporterToutesLesFeuilles = swDxfMultisheet_e.swDxfActiveSheetOnly;
                    if (ToutesLesFeuilles)
                        Sw.DxfDwg_ExporterToutesLesFeuilles = swDxfMultisheet_e.swDxfMultiSheet;
                    break;
                case eTypeFichierExport.DWG:
                    Sw.DxfDwg_ExporterToutesLesFeuilles = swDxfMultisheet_e.swDxfActiveSheetOnly;
                    if (ToutesLesFeuilles)
                        Sw.DxfDwg_ExporterToutesLesFeuilles = swDxfMultisheet_e.swDxfMultiSheet;
                    break;
                case eTypeFichierExport.PDF:
                    OptionsPDF = App.Sw.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);
                    if (ToutesLesFeuilles)
                    {
                        OptionsPDF.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportAllSheets, null);
                    }
                    else
                    {
                        DispatchWrapper[] Wrapper = new DispatchWrapper[1];
                        Wrapper[0] = new DispatchWrapper((Object)feuille);
                        OptionsPDF.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, Wrapper);
                    }
                    break;
            }

            String CheminFichier = feuille.GetName();

            int Erreur = 0;
            int Warning = 0;

            if (!String.IsNullOrEmpty(NomDuFichierAlternatif))
                CheminFichier = NomDuFichierAlternatif;

            CheminFichier = Path.Combine(CheminDossier, CheminFichier + TypeExport.GetEnumInfo<ExtFichier>());
            dessin.eModelDoc2().Extension.SaveAs(CheminFichier, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, OptionsPDF, Erreur, Warning);

            return CheminFichier;
        }

        public static String eMettreEnPagePourImpression(this DrawingDoc dessin, Sheet feuille, swPageSetupDrawingColor_e couleur = swPageSetupDrawingColor_e.swPageSetup_AutomaticDrawingColor, Boolean hauteQualite = false)
        {
            ModelDoc2 mdl = dessin.eModelDoc2();
            ModelDocExtension ext = dessin.eModelDoc2().Extension;

            eFormat format = feuille.eGetFormat();
            eOrientation orientation = feuille.eGetOrientation();

            ext.UsePageSetup = (int)swPageSetupInUse_e.swPageSetupInUse_DrawingSheet;
            PageSetup pSetupFeuille = feuille.PageSetup;
            pSetupFeuille.DrawingColor = (int)couleur;
            pSetupFeuille.HighQuality = hauteQualite;
            //pSetupFeuille.PrinterPaperSource = 15;
            pSetupFeuille.PrinterPaperSize = (int)format;
            pSetupFeuille.ScaleToFit = false;

            if (orientation == eOrientation.Paysage)
                pSetupFeuille.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
            else
                pSetupFeuille.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Portrait;

            return format.GetEnumInfo<Intitule>() + " " + orientation.GetEnumInfo<Intitule>();
        }

        public static Sheet eFeuilleActive(this DrawingDoc dessin)
        {
            return dessin.GetCurrentSheet();
        }

        public static List<Sheet> eListeDesFeuilles(this DrawingDoc dessin)
        {
            List<Sheet> pListeFeuilles = new List<Sheet>();

            if (dessin.GetSheetCount() == 0)
                return pListeFeuilles;

            foreach (String NomFeuille in dessin.GetSheetNames())
            {
                Sheet pSwFeuille = dessin.get_Sheet(NomFeuille);
                pListeFeuilles.Add(pSwFeuille);
            }

            return pListeFeuilles;
        }

        public static void eParcourirLesFeuilles(this DrawingDoc dessin, Predicate<Sheet> filtre)
        {
            if (dessin.GetSheetCount() == 0)
                return;

            foreach (String NomFeuille in dessin.GetSheetNames())
            {
                if (filtre(dessin.get_Sheet(NomFeuille)))
                    break;
            }
        }
    }

    public abstract class eGeomBase
    {
        public Boolean Maj = true;

        public delegate void OnModifyEventHandler();

        public event OnModifyEventHandler OnModify;

        public void Modify()
        {
            if (!Maj && OnModify.IsRef())
                OnModify();
        }
    }

    public class ePoint : eGeomBase
    {
        private Double _X = 0;
        private Double _Y = 0;
        private Double _Z = 0;

        public ePoint() { }

        public ePoint(Double x, double y, Double z) { X = x; Y = y; Z = z; Maj = false; }

        public Double X { get { return _X; } set { _X = value; Modify(); } }
        public Double Y { get { return _Y; } set { _Y = value; Modify(); } }
        public Double Z { get { return _Z; } set { _Z = value; Modify(); } }

        public void Deplacer(eVecteur V)
        {
            X += V.X; Y += V.Y; Z += V.Z;
        }

        public void Echelle(Double S)
        {
            X *= S; Y *= S; Z *= S;
        }

        public ePoint Additionner(eVecteur V)
        {
            return new ePoint(X + V.X, Y + V.Y, Z + V.Z);
        }

        public ePoint Multiplier(Double S)
        {
            return new ePoint(X * S, Y * S, Z * S);
        }
    }

    public class eVecteur
    {
        public eVecteur(Double X, Double Y, Double Z) { this.X = X; this.Y = Y; this.Z = Z; }

        public Double X { get; set; }
        public Double Y { get; set; }
        public Double Z { get; set; }
    }

    public class eRepere
    {
        private ePoint _Origine = new ePoint(0, 0, 0);
        private eVecteur _VecteurX = new eVecteur(0, 0, 0);
        private eVecteur _VecteurY = new eVecteur(0, 0, 0);
        private eVecteur _VecteurZ = new eVecteur(0, 0, 0);

        public ePoint Origine { get { return _Origine; } set { _Origine = value; } }
        public eVecteur VecteurX { get { return _VecteurX; } set { _VecteurX = value; } }
        public eVecteur VecteurY { get { return _VecteurY; } set { _VecteurY = value; } }
        public eVecteur VecteurZ { get { return _VecteurZ; } set { _VecteurZ = value; } }
        public Double Echelle { get; set; }
    }

    public class eRectangle : eGeomBase
    {
        private Double _Lg = 0;
        private Double _Ht = 0;

        public Double Lg { get { return _Lg; } set { _Lg = value; Modify(); } }
        public Double Ht { get { return _Ht; } set { _Ht = value; Modify(); } }

        public eRectangle(Double lg, Double ht) { Lg = lg; Ht = ht; Maj = false; }
    }

    public class eZone
    {
        private ePoint _Centre = new ePoint(0, 0, 0);

        private eRectangle _Rectangle = new eRectangle(0, 0);

        private ePoint _PointMin = new ePoint(0, 0, 0);
        private ePoint _PointMax = new ePoint(0, 0, 0);

        public ePoint PointMin { get { return _PointMin; } }
        public ePoint PointMax { get { return _PointMax; } }

        public eZone()
        { }

        private void Maj()
        {
            _Centre.X = (_PointMin.X + PointMax.X) * 0.5;
            _Centre.Y = (_PointMin.Y + PointMax.Y) * 0.5;
            _Centre.Z = (_PointMin.Z + PointMax.Z) * 0.5;

            _Rectangle.Lg = _PointMax.X - _PointMin.X;
            _Rectangle.Ht = _PointMax.Y - _PointMin.Y;
        }

        public ePoint CentreZone { get { Maj(); return _Centre; } }

        public eRectangle Rectangle { get { Maj(); return _Rectangle; } }
    }
}
