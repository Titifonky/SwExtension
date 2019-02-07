using LogDebugging;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;

namespace ModuleProduction.ModuleRepereCorps
{
    public class CmdRepereCorps : Cmd
    {
        public ModelDoc2 MdlBase = null;

        protected override void Command()
        {
            try
            { }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }
}


