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
        ModuleTitre("Test6"),
        ModuleNom("Test6")]
    public class Test6 : BoutonBase
    {
        public Test6() { }

        protected override void Command()
        {
            try
            {
                if (MdlBase.eSelect_RecupererTypeObjet() != e_swSelectType.swSelFACES)
                    return;

                var face = MdlBase.eSelect_RecupererObjet<Face2>();

                MdlBase.eEffacerSelection();

                var corps = (Body2)face.GetBody();

                Lancer(corps);

            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }

        private void Lancer(Body2 corps)
        {
            Face2 faceBase = null;
            var listeFunction = (object[])corps.GetFeatures();

            foreach (Feature feature in listeFunction)
            {
                WindowLog.Ecrire(feature.GetTypeName2());
                if (feature.GetTypeName2() == FeatureType.swTnFlatPattern)
                {
                    var def = (FlatPatternFeatureData)feature.GetDefinition();
                    faceBase = (Face2)def.FixedFace2;
                    break;
                }
            }

            if (faceBase.IsNull()) return;

            var listeLoop = new List<Loop2>();

            foreach (Loop2 loop in faceBase.GetLoops())
                if (!loop.IsOuter())
                    listeLoop.Add(loop);

            foreach (var loop in listeLoop)
            {
                var edge = (Edge)loop.GetEdges()[0];

                Face2 faceCylindre = null;
                foreach (Face2 face in edge.GetTwoAdjacentFaces2())
                {
                    if (!face.IsSame(faceBase))
                    {
                        faceCylindre = face;
                        break;
                    }
                }


            }
        }
    }
}
