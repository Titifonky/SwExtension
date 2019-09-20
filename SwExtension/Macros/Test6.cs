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
                var comp = MdlBase.eSelect_RecupererComposant();

                FaceBase = face;
                MdlBase.eEffacerSelection();

                WindowLog.Ecrire("Recherche du corps du composant");

                var corps = (Body2)face.GetBody();

                WindowLog.Ecrire(corps.Name);

                corps = comp.eChercherCorps(corps.Name, false);

                WindowLog.Ecrire(corps.Name);

                Lancer(corps, comp);

            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }

        private void Lancer(Body2 corps, Component2 comp)
        {
            Face2 faceBase = null;
            var listeFunction = (object[])corps.GetFeatures();

            // On recherche la face de base de la tôle
            // On part de la fonction dépliée pour récupérer la face fixe.
            foreach (Feature feature in (object[])corps.GetFeatures())
            {
                if (feature.GetTypeName2() == FeatureType.swTnFlatPattern)
                {
                    // On récupère la face fixe
                    var def = (FlatPatternFeatureData)feature.GetDefinition();
                    faceBase = (Face2)def.FixedFace2;

                    // On liste les faces du corps et on regarde 
                    // la face qui est la même que celle de la face fixe
                    foreach (var face in corps.eListeDesFaces())
                    {
                        if(face.IsSame(faceBase))
                        {
                            faceBase = face;
                            break;
                        }
                    }
                    break;
                }
            }

            if (faceBase.IsNull())
            {
                WindowLog.Ecrire("Pas de face de base");
                return;
            }

            MdlBase.eEffacerSelection();

            var listeLoop = (Object[])faceBase.GetLoops();
            WindowLog.Ecrire("Recherche des loop // Nb de loop : " + listeLoop.Length);

            var listePercage = new List<Loop2>();

            foreach (Loop2 loop in listeLoop)
                if (!loop.IsOuter())
                    listePercage.Add(loop);

            MdlBase.eEffacerSelection();

            WindowLog.Ecrire("Selection des percages // Nb de percage : " + listePercage.Count);
            foreach (var loop in listePercage)
            {
                var edge = (Edge)loop.GetEdges()[0];

                Face2 faceCylindre = null;
                foreach (Face2 face in edge.GetTwoAdjacentFaces2())
                {
                    if (!face.IsSame(faceBase))
                    {
                        faceCylindre = face;
                        faceCylindre.eSelectEntite(MdlBase, -1, true);
                        break;
                    }
                }
            }
        }
    }
}
