using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace ModuleCreerSymetrie
{
    public class CmdCreerSymetrie : Cmd
    {
        public ModelDoc2 MdlBase = null;
        public Feature Plan = null;
        public List<Body2> ListeCorps = null;

        protected override void Command()
        {
            SymetriserPiece(MdlBase, Plan, ListeCorps);
        }

        private void SymetriserPiece(ModelDoc2 mdl , Feature plan, List<Body2> listeCorps)
        {
            try
            {
                // Création de la symetrie
                mdl.eEffacerSelection();
                mdl.eSelectMulti(plan, 2, true);
                mdl.eSelectMulti(listeCorps, 256, true);
                var Symetrie = mdl.FeatureManager.InsertMirrorFeature2(true, false, false, false, (int)swFeatureScope_e.swFeatureScope_AllBodies);

                // Suppression des pièces symétrisée
                mdl.eEffacerSelection();
                mdl.eSelectMulti(listeCorps, -1, false);
                var Supprimer = mdl.FeatureManager.InsertDeleteBody2(false);

                // On met le tout dans un dossier correctement renommé
                mdl.eEffacerSelection();
                mdl.EditRebuild3();
                Symetrie.eSelect();
                Supprimer.eSelect(true);
                Feature Dossier = mdl.FeatureManager.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing);
                Dossier.eRenommerFonction(mdl.eNomConfigActive() + "-Symetrie");

                // On nettoie
                mdl.eEffacerSelection();
                mdl.EditRebuild3();
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }
    }
}