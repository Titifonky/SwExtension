using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Piece),
        ModuleTitre("Test4"),
        ModuleNom("Test4")]
    public class Test4 : BoutonBase
    {
        private readonly Double Decal = 40;

        public Test4() { }

        protected override void Command()
        {
            try
            {

                if (MdlBase.eSelect_RecupererTypeObjet() != e_swSelectType.swSelFACES)
                    return;

                var face = MdlBase.eSelect_RecupererObjet<Face2>();

                var sm = MdlBase.SketchManager;

                sm.InsertSketch(false);

                MdlBase.SketchOffsetEntities2(Decal * -0.001, false, false);

                MdlBase.eEffacerSelection();

                var sk = sm.ActiveSketch;

                SketchPoint PremierPoint = null;
                var arrCt = (object[])sk.GetSketchContours();
                foreach (SketchContour ct in arrCt)
                {
                    var arrSeg = (object[])ct.GetSketchSegments();
                    foreach (SketchSegment sg in arrSeg)
                    {
                        if (sg.eType() != swSketchSegments_e.swSketchLINE) continue;

                        sg.ConstructionGeometry = true;
                        var sl = (SketchLine)sg;
                        var lg = sg.GetLength() * 1000.0;
                        var nb = Math.Floor(lg / 300.0);
                        var pas = lg / (nb + 1);
                        nb = Math.Max(0, pas > 200 ? nb : nb - 1);
                        var ptDepart = new ePoint((SketchPoint)sl.GetStartPoint2());
                        sm.CreatePoint(ptDepart.X, ptDepart.Y, ptDepart.Z);
                        if (nb > 0)
                        {
                            sg.eSelect(MdlBase, 1, false);
                            sg.EqualSegment((int)swSketchSegmentType_e.swSketchSegmentType_sketchpoints, (int)nb);
                        }

                        if(PremierPoint.IsNull())
                        {
                            PremierPoint = (SketchPoint)sl.GetStartPoint2();

                            sm.AddToDB = true;

                            var sgCercle = sm.CreateCircleByRadius(PremierPoint.X, PremierPoint.Y, PremierPoint.Z, 0.005);
                            PremierPoint.eSelect(MdlBase, 0, false);
                            sgCercle.eSelect(MdlBase, 0, true);
                            MdlBase.SketchAddConstraints("sgCONCENTRIC");

                            sgCercle.eSelect(MdlBase, 0, false);

                            var reg = App.Sw.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
                            App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

                            var dispdim = (DisplayDimension)MdlBase.AddDiameterDimension2(PremierPoint.X, PremierPoint.Y, PremierPoint.Z);

                            App.Sw.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, reg);
                            //var dim = dispdim.GetDimension2(0);

                            sm.AddToDB = false;
                        }
                    }
                }

                sm.InsertSketch(true);

                var fm = MdlBase.FeatureManager;

                MdlBase.eEffacerSelection();

                var bd = (Body2)face.GetBody();
                var fSketch = (Feature)sk;

                fSketch.eSelectionnerById2(MdlBase, 0, true);
                bd.eSelect(MdlBase, 8, true);
                var fPercage = fm.FeatureCut4(true, false, false, (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind, 0.02, 0.02, false, false, false, false, 0, 0, false, false, false, false, false, true, false, true, true, false, 0, 0, false, false);

                MdlBase.eEffacerSelection();

                fPercage.eSelectionnerById2(MdlBase, 4, true);
                PremierPoint.eSelect(MdlBase, 32, true);
                fSketch.eSelectionnerById2(MdlBase, 64, true);
                bd.eSelect(MdlBase, 512, true);

                var featRepetDef = (SketchPatternFeatureData)fm.CreateDefinition((int)swFeatureNameID_e.swFmSketchPattern);
                featRepetDef.GeometryPattern = true;
                featRepetDef.UseCentroid = false;
                var featRepet = fm.CreateFeature(featRepetDef);

                WindowLog.Ecrire("Repetition : " + featRepet.IsRef());

            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }
    }
}
