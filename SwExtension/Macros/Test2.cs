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
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Test2"),
        ModuleNom("Test2")]
    public class Test2 : BoutonBase
    {
        private Parametre Pid;

        public Test2()
        {
            _Config = new ConfigModule(typeof(Test));
            Pid = _Config.AjouterParam("Pid", "xxx");
        }

        protected override void Command()
        {
            try
            {
                Byte[] Tab = File.ReadAllBytes(Path.Combine(MdlBase.eDossier(), "Corps.data"));
                MemoryStream ms = new MemoryStream(Tab);
                ManagedIStream MgIs = new ManagedIStream(ms);
                Modeler mdlr = (Modeler)App.Sw.GetModeler();
                var corps = (Body2)mdlr.Restore(MgIs);
                
                var err = corps.Check3;
                WindowLog.EcrireF("nb erreurs : {0}", err.Count);

                var retval = corps.Display3(MdlBase, 255, (int)swTempBodySelectOptions_e.swTempBodySelectable);
                WindowLog.Ecrire("Temporary body displayed (0 = success)? " + retval);
            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }
    }
}
