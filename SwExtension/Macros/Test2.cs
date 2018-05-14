using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
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

                //var F = mdl.eSelect_RecupererObjet<Feature>(1);
                //mdl.eEffacerSelection();

                //var def = (StructuralMemberFeatureData)F.GetDefinition();
                //WindowLog.Ecrire(def.WeldmentProfilePath);
                //WindowLog.Ecrire(def.ConfigurationName);
                //foreach (var sf in F.eListeSousFonction())
                //{
                //    WindowLog.Ecrire(sf.GetTypeName2());
                //}

                var Face = mdl.eSelect_RecupererObjet<Face2>(1);
                Body2 Corps = null;
                if (Face.IsRef())
                    Corps = Face.GetBody();
                else
                    Corps = mdl.eSelect_RecupererObjet<Body2>(1);

                mdl.eEffacerSelection();

                var SM = mdl.SketchManager;

                SM.Insert3DSketch(false);
                SM.AddToDB = true;
                SM.DisplayWhenAdded = false;

                List<Vecteur> ListeDir = new List<Vecteur>();
                ListeDir.Add(new Vecteur(1, 0, 0));
                ListeDir.Add(new Vecteur(-1, 0, 0));
                ListeDir.Add(new Vecteur(0, 1, 0));
                ListeDir.Add(new Vecteur(0, -1, 0));
                ListeDir.Add(new Vecteur(0, 0, 1));
                ListeDir.Add(new Vecteur(0, 0, -1));

                foreach (var v in ListeDir)
                {
                    var Pt = Corps.ePointExtreme(v);
                    SM.CreatePoint(Pt.X, Pt.Y, Pt.Z);
                }

                //{
                //    foreach (var Face in corps.eListeDesFaces())
                //    {
                //        var S = (Surface)Face.GetSurface();

                //        var UV = (Double[])Face.GetUVBounds();

                //        Boolean Reverse = Face.FaceInSurfaceSense();

                //        var ev1 = (Double[])S.Evaluate((UV[0] + UV[1]) * 0.5, (UV[2] + UV[3]) * 0.5, 0, 0);

                //        Vecteur Vn = new Vecteur(ev1[3], ev1[4], ev1[5]);
                //        if (Reverse)
                //            Vn.Inverser();

                //        Vn.Normaliser();
                //        Vn.Multiplier(0.01);

                //        Point Pt = new Point();

                //        if (S.IsPlane())
                //        {
                //            Double[] Param = S.PlaneParams;

                //            Pt = new Point(Param[3], Param[4], Param[5]);
                //        }
                //        else if (S.IsCylinder())
                //        {
                //            Double[] Param = S.CylinderParams;

                //            Pt = new Point(Param[0], Param[1], Param[2]);
                //        }
                //        else
                //            continue;

                //        SM.CreatePoint(Pt.X, Pt.Y, Pt.Z);
                //        SM.CreateLine(Pt.X, Pt.Y, Pt.Z, Pt.X + Vn.X, Pt.Y + Vn.Y, Pt.Z + Vn.Z);
                //    }
                //}

                SM.DisplayWhenAdded = true;
                SM.AddToDB = false;
                SM.Insert3DSketch(true);



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



    }
}
