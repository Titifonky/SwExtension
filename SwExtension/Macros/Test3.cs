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
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Test3"),
        ModuleNom("Test3")]
    public class Test3 : BoutonBase
    {
        private readonly Double Decal = 40;

        public Test3() { }

        protected override void Command()
        {
            try
            {
                Face2 face = null;
                SketchSegment segment = null;

                if (MdlBase.eSelect_RecupererTypeObjet(1) == e_swSelectType.swSelFACES)
                {
                    face = MdlBase.eSelect_RecupererObjet<Face2>(1);
                    segment = MdlBase.eSelect_RecupererObjet<SketchSegment>(2);
                }
                else
                {
                    face = MdlBase.eSelect_RecupererObjet<Face2>(2);
                    segment = MdlBase.eSelect_RecupererObjet<SketchSegment>(1);
                }

                if (face == null || segment == null) return;

                MdlBase.eEffacerSelection();

                Boolean r = false;
                Boolean reverse = false;

                var sk = segment.GetSketch();
                var xform = (MathTransform)sk.ModelToSketchTransform.Inverse();

                if (segment.GetType() != (int)swSketchSegments_e.swSketchLINE) return;

                var sl = (SketchLine)segment;

                var start = new ePoint(sl.GetStartPoint2());
                var end = new ePoint(sl.GetEndPoint2());

                start.ApplyMathTransform(xform);
                end.ApplyMathTransform(xform);

                WindowLog.Ecrire(start.IsRef() + " " + start.ToString());
                WindowLog.Ecrire(end.IsRef() + " " + end.ToString());

                var box = (Double[])face.GetBox();

                var pt = new ePoint((box[3] + box[0]) * 0.5, (box[4] + box[1]) * 0.5, (box[5] + box[2]) * 0.5);
                WindowLog.Ecrire(pt.IsRef() + " " + pt.ToString());

                if (start.Distance2(pt) > end.Distance2(pt))
                    reverse = true;

                r = face.eSelectEntite(MdlBase, 4, false);

                r = segment.eSelect(MdlBase, 1, true);

                var cp = (Body2)face.GetBody();
                r = cp.eSelect(MdlBase, 512, true);

                var fm = MdlBase.FeatureManager;
                var featRepetDef = (CurveDrivenPatternFeatureData)fm.CreateDefinition((int)swFeatureNameID_e.swFmCurvePattern);

                featRepetDef.D1AlignmentMethod = 0;
                featRepetDef.D1CurveMethod = 0;
                featRepetDef.D1InstanceCount = 3;
                featRepetDef.D1IsEqualSpaced = true;
                featRepetDef.D1ReverseDirection = reverse;
                featRepetDef.D1Spacing = 0.001;
                featRepetDef.D2InstanceCount = 1;
                featRepetDef.D2IsEqualSpaced = false;
                featRepetDef.D2PatternSeedOnly = false;
                featRepetDef.D2ReverseDirection = false;
                featRepetDef.D2Spacing = 0.001;
                featRepetDef.GeometryPattern = true;

                var featRepet = fm.CreateFeature(featRepetDef);

                WindowLog.Ecrire(featRepet != null);
            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
                WindowLog.Ecrire(new Object[] { e });
            }
        }
    }
}
