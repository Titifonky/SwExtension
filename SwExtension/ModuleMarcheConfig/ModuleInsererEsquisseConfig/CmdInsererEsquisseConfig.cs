using SolidWorks.Interop.sldworks;
using SwExtension;
using System;

namespace ModuleMarcheConfig
{
    namespace ModuleInsererEsquisseConfig
    {
        public class CmdInsererEsquisseConfig : Cmd
        {
            public ModelDoc2 MdlBase = null;
            public Feature Plan = null;
            public String NomEsquisse = "Config";

            protected override void Command()
            {
                String TypeSel = "";
                MdlBase.Extension.SelectByID2(Plan.GetNameForSelection(out TypeSel), TypeSel, 0, 0, 0, false, -1, null, 0);

                SketchManager Sk = MdlBase.SketchManager;

                Sk.InsertSketch(true);


                SketchSegment F = Sk.CreateLine(0, 0, 0, -0.064085, 0.171639, 0);

                Sk.InsertSketch(true);

                MdlBase.EditRebuild3();
            }
        }
    }
}


