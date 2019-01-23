using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using TriangleNet;
using TriangleNet.Geometry;
using TriangleNet.Meshing;
using TriangleNet.Voronoi;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Decompter les perçages"),
        ModuleNom("DecompterPercage")]
    public class BoutonDecompterPercage : BoutonBase
    {

        private Dictionary<Double, int> DicQte = new Dictionary<Double, int>();

        protected override void Command()
        {
            try
            {
                if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                    Decompter(MdlBase.eComposantRacine());
                else
                    MdlBase.eRecParcourirComposants(Decompter);

                WindowLog.SautDeLigne();

                foreach (var item in DicQte)
                    WindowLog.EcrireF("{0,-5} : {1,-5}", "Ø " + Math.Round(item.Key, 2), "×" + item.Value);

            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }

        private Boolean Decompter(Component2 cp)
        {
            try
            {
                if (!cp.IsHidden(true))
                {
                    foreach (var corps in cp.eListeCorps())
                    {
                        foreach (var face in corps.eListeDesFaces())
                        {
                            Surface S = face.GetSurface();
                            if (S.IsRef() && S.IsCylinder() && (face.GetLoopCount() > 1))
                            {
                                Double[] ListeParam = (Double[])S.CylinderParams;
                                Double Diam = Math.Round(ListeParam[6] * 2.0 * 1000, 2);

                                DicQte.AddIfNotExistOrPlus(Diam);
                            }
                        }
                    }
                }

            }
            catch (Exception e) { this.LogMethode(new Object[] { e }); }

            return false;
        }
    }
}
