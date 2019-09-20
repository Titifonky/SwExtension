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

        private Face2 FaceBase = null;

        protected override void Command()
        {
            try
            {
                if (MdlBase.eSelect_RecupererTypeObjet() != e_swSelectType.swSelFACES)
                    return;

                var face = MdlBase.eSelect_RecupererObjet<Face2>();

                FaceBase = face;
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
            MdlBase.eEffacerSelection();

            var faceBase = eFaceFixe(corps);

            var liste = new List<Face2>();

            eFacesTangentes(faceBase, ref liste);

            foreach (var face in liste)
                face.eSelectEntite(MdlBase, -1, true);

            //if (faceBase.IsNull())
            //{
            //    WindowLog.Ecrire("Pas de face de base");
            //    return;
            //}

            //MdlBase.eEffacerSelection();

            //var listeLoop = (Object[])faceBase.GetLoops();
            //WindowLog.Ecrire("Recherche des loop // Nb de loop : " + listeLoop.Length);

            //var listePercage = new List<Loop2>();

            //foreach (Loop2 loop in listeLoop)
            //    if (!loop.IsOuter())
            //        listePercage.Add(loop);

            //MdlBase.eEffacerSelection();

            //WindowLog.Ecrire("Selection des percages // Nb de percage : " + listePercage.Count);
            //foreach (var loop in listePercage)
            //{
            //    var edge = (Edge)loop.GetEdges()[0];

            //    Face2 faceCylindre = null;
            //    foreach (Face2 face in edge.GetTwoAdjacentFaces2())
            //    {
            //        if (!face.IsSame(faceBase))
            //        {
            //            faceCylindre = face;
            //            faceCylindre.eSelectEntite(MdlBase, -1, true);
            //            break;
            //        }
            //    }
            //}
        }

        private Face2 eFaceFixe(Body2 tole)
        {
            Face2 faceFixe = null;

            // On recherche la face de base de la tôle
            // On part de la fonction dépliée pour récupérer la face fixe.
            foreach (Feature feature in (object[])tole.GetFeatures())
            {
                if (feature.GetTypeName2() == FeatureType.swTnFlatPattern)
                {
                    // On récupère la face fixe
                    var def = (FlatPatternFeatureData)feature.GetDefinition();
                    faceFixe = (Face2)def.FixedFace2;

                    // On liste les faces du corps et on regarde 
                    // la face qui est la même que celle de la face fixe
                    foreach (var face in tole.eListeDesFaces())
                    {
                        if (face.IsSame(faceFixe))
                        {
                            faceFixe = face;
                            break;
                        }
                    }
                    break;
                }
            }

            return faceFixe;
        }

        private void eFacesTangentes(Face2 face, ref List<Face2> listeFaces)
        {
            var listeFacesTangentes = new List<Face2>();

            foreach (Loop2 loop in (Object[])face.GetLoops())
            {
                foreach (Edge edge in loop.GetEdges())
                {
                    if(eFacesContiguesTangentes(edge))
                    {
                        foreach (var faceTangente in edge.eListeDesFaces())
                        {
                            Boolean ajouter = true;
                            foreach (var f in listeFaces)
                            {
                                if (f.IsSame(faceTangente))
                                {
                                    ajouter = false;
                                    break;
                                }
                            }

                            if(ajouter)
                            {
                                listeFaces.Add(faceTangente);
                                listeFacesTangentes.Add(faceTangente);
                            }
                        }
                    }
                }
            }

            foreach (var faceTangente in listeFacesTangentes)
                eFacesTangentes(faceTangente, ref listeFaces);
        }

        private Boolean eFacesContiguesTangentes(Edge edge)
        {
            var pt = ePointMilieu(edge);
            var lf = edge.eListeDesFaces();
            var f1 = (Surface)lf[0].GetSurface();
            var f2 = (Surface)lf[1].GetSurface();
            var v1 = new eVecteur((double[])f1.EvaluateAtPoint(pt.X, pt.Y, pt.Z));
            var v2 = new eVecteur((double[])f2.EvaluateAtPoint(pt.X, pt.Y, pt.Z));

            if (v1.EstColineaire(v2, 1E-10, false))
                return true;

            return false;
        }

        public ePoint ePointMilieu(Edge edge)
        {
            edge.GetCurve();
            Curve Courbe = edge.GetCurve();
            var param = (CurveParamData)edge.GetCurveParams3();
            var start = param.UMinValue;
            var end = param.UMaxValue;
            if (!param.Sense)
            {
                var t = end * -1;
                end = start * -1;
                start = t;
            }

            var r = (double[])edge.Evaluate2((end + start) * 0.5, 0);

            return new ePoint(r);
        }
    }
}
