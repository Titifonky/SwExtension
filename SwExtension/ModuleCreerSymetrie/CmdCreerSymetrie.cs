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
            try
            {
                var Symetrie = MdlBase.FeatureManager.InsertMirrorFeature2(true, false, false, false, (int)swFeatureScope_e.swFeatureScope_AllBodies);
                MdlBase.eEffacerSelection();
                MdlBase.eSelectMulti(ListeCorps, -1, false);
                var Supprimer = MdlBase.FeatureManager.InsertDeleteBody2(false);
                MdlBase.eEffacerSelection();
                MdlBase.EditRebuild3();

                Symetrie.eSelect();
                Supprimer.eSelect(true);
                Feature Dossier = MdlBase.FeatureManager.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing);
                Dossier.eRenommerFonction(MdlBase.eNomConfigActive() + "-Symetrie");

                MdlBase.eEffacerSelection();
                MdlBase.EditRebuild3();
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }
    }
}