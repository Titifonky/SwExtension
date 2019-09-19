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
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Test2"),
        ModuleNom("Test2")]
    public class Test2 : BoutonBase
    {
        private readonly Double Decal = 40;

        public Test2() { }

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

                var arr = (object[])sk.GetSketchSegments();
                foreach (SketchSegment sg in arr)
                    sg.ConstructionGeometry = true;

                sm.InsertSketch(true);

                MdlBase.eEffacerSelection();

            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }
    }
}
