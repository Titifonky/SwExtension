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
        ModuleTitre("Test7"),
        ModuleNom("Test7")]
    public class Test7 : BoutonBase
    {
        public Test7() { }

        protected override void Command()
        {
            try
            {
                try
                {
                    

                }
                catch (Exception e)
                {
                    this.LogMethode(new Object[] { e });
                }

            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }
    }
}
