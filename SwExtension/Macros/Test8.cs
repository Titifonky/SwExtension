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
    [ModuleTypeDocContexte(eTypeDoc.Piece),
        ModuleTitre("Test8"),
        ModuleNom("Test8")]
    public class Test8 : BoutonBase
    {
        public Test8() { }

        protected override void Command()
        {
            try
            {
                var mass = MdlBase.Extension.CreateMassProperty();

                WindowLog.EcrireF("Ecraser la masse : {0}", mass.OverrideMass);

                mass.OverrideMass = false;


            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }
    }
}
