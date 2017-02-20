using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;

namespace ModuleListerPercage
{
    public class CmdListerPercage : Cmd
    {
        public ModelDoc2 MdlBase = null;

        protected override void Command()
        {
            try
            { }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }
    }
}


