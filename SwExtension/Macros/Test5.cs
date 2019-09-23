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
        ModuleTitre("Test5"),
        ModuleNom("Test5")]
    public class Test5 : BoutonBase
    {
        public Test5() { }

        protected override void Command()
        {
            try
            {
                if (MdlBase.eSelect_RecupererTypeObjet() != e_swSelectType.swSelFACES)
                    return;

                var liste = MdlBase.eSelect_RecupererListeObjets<Face2>();

                MdlBase.eEffacerSelection();

                foreach (var face in liste)
                    Decaler(face, 0.0005, true);

            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }

        public void Decaler(Face2 face, Double decal, Boolean inverser)
        {
            Loop2 bord = null;

            foreach (Loop2 loop in face.GetLoops())
                if (loop.IsOuter())
                {
                    bord = loop;
                    break;
                }

            var listeFace = new List<Face2>();
            foreach (Edge e in bord.GetEdges())
                listeFace.Add(e.eAutreFace(face));

            MdlBase.eEffacerSelection();

            foreach (var f in listeFace)
                f.eSelectEntite(MdlBase, 1, true);

            var feat = MdlBase.FeatureManager.InsertMoveFace3((int)swMoveFaceType_e.swMoveFaceTypeOffset, inverser, 0, decal, null, null, 0, 0);

            WindowLog.Ecrire("Décalage : " + feat.IsRef());
        }
    }
}
