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
        ModuleTitre("Test2"),
        ModuleNom("Test2")]

    public class Test2 : BoutonBase
    {
        protected override void Command()
        {
            try
            {
                ModelDoc2 mdl = App.ModelDoc2;
                //var DossierExport = mdl.eDossier();
                //var NomFichier = mdl.eNomSansExt();

                var Liste = mdl.eSelect_RecupererListeObjets<Face2>();
                var Face1 = Liste[0];
                var Face2 = Liste[1];
                //Body2 Corps = Face.GetBody();
                mdl.eEffacerSelection();

                Byte[] id1 = mdl.Extension.GetPersistReference(Face1);
                Byte[] id2 = mdl.Extension.GetPersistReference(Face2);
                String s = "";
                for (int i = 0; i < Math.Max(id1.Length, id2.Length); i++)
                {
                    if (i < id1.Length && i < id2.Length)
                    {
                        if (id1[i] == id2[i])
                            s += "|";
                        else
                            s += "[" + id1[i] + "|" + id2[i] + "]";
                    }
                    else if (i < id1.Length)
                        s += "1[" + id1[i] + "]";
                    else if (i < id2.Length)
                        s += "2[" + id2[i] + "]";
                }

                WindowLog.Ecrire(s);

                //var SM = mdl.SketchManager;

                //var UV = (Double[])Face.GetUVBounds();

                //var S = (Surface)Face.GetSurface();

                //Boolean Reverse = Face.FaceInSurfaceSense();

                //var ev1 = (Double[])S.Evaluate((UV[0] + UV[1]) * 0.5, (UV[2] + UV[3]) * 0.5, 0, 0);
                //if (Reverse)
                //{
                //    ev1[3] = -ev1[3];
                //    ev1[4] = -ev1[4];
                //    ev1[5] = -ev1[5];
                //}

                //SM.Insert3DSketch(false);
                //SM.AddToDB = true;
                //SM.DisplayWhenAdded = false;

                //SM.CreatePoint(ev1[0], ev1[1], ev1[2]);
                //SM.CreateLine(ev1[0], ev1[1], ev1[2], ev1[0] + ev1[3], ev1[1] + ev1[4], ev1[2] + ev1[5]);

                //SM.DisplayWhenAdded = true;
                //SM.AddToDB = false;
                //SM.Insert3DSketch(true);

            }
            catch (Exception e) { this.LogMethode(new Object[] { e }); }

        }

        private Double[] NormaleCylindre(Double[] Axe)
        {
            Double[] Normale = new Double[] { 0, 0, 0 };

            if (Axe[0] == 0 && Axe[1] == 0)
                Normale[0] = 1;
            else
            {
                Normale[0] = Axe[1];
                Normale[1] = -1 * Axe[0];
            }

            return Normale;
        }
    }
}
