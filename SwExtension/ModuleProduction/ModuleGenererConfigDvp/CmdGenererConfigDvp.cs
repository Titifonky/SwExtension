using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace ModuleProduction.ModuleGenererConfigDvp
{
    public class CmdGenererConfigDvp : Cmd
    {
        public ModelDoc2 MdlBase
        {
            get { return _mdlBase; }
            set { _mdlBase = value; }
        }

        private ModelDoc2 _mdlBase = null;

        public Boolean SupprimerLesAnciennesConfigs = false;

        protected override void Command()
        {
            try
            {
                var ListeCorps = MdlBase.pChargerNomenclature();

                foreach (var corps in ListeCorps.Values)
                {
                    WindowLog.EcrireF("{0} -> dvp", corps.Repere);
                    corps.pCreerDvp(MdlBase.pDossierPiece(), SupprimerLesAnciennesConfigs);
                }

                MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                MdlBase.EditRebuild3();
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }
}


