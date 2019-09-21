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

                var faceSel = MdlBase.eSelect_RecupererObjet<Face2>();

                MdlBase.eEffacerSelection();

                var faceBase = faceSel;

                var liste = new List<Face2>();

                faceBase.eChercherFacesTangentes(ref liste);

                //liste = liste.FindAll(f => ((Surface)f.GetSurface()).IsPlane());

                foreach (var face in liste)
                    face.eSelectEntite(MdlBase, -1, true);

            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }

        private void Lancer(Body2 corps)
        {
            MdlBase.eEffacerSelection();

            var faceBase = corps.eFaceFixeTolerie();

            var liste = new List<Face2>();

            faceBase.eChercherFacesTangentes(ref liste);

            liste = liste.FindAll(f => ((Surface)f.GetSurface()).IsPlane());

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
    }
}
