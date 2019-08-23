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
                //Body2 CorpsBase = null;
                //if (MdlBase.eSelect_RecupererTypeObjet() == e_swSelectType.swSelFACES)
                //{
                //    var face = MdlBase.eSelect_RecupererObjet<Face2>();
                //    CorpsBase = face.GetBody();
                //}
                //else if (MdlBase.eSelect_RecupererTypeObjet() == e_swSelectType.swSelSOLIDBODIES)
                //{
                //    CorpsBase = MdlBase.eSelect_RecupererObjet<Body2>();
                //}

                //if (CorpsBase == null)
                //{
                //    WindowLog.Ecrire("Erreur de corps selectionné");
                //    return;
                //}


                //WindowLog.Ecrire("Type de corps : " + CorpsBase.eTypeDeCorps());

                if(MdlBase.eSelect_RecupererTypeObjet() == e_swSelectType.swSelBODYFEATURES)
                {
                    var f = MdlBase.eSelect_RecupererObjet<Feature>();
                    WindowLog.Ecrire("Fonction : " + f.Name);
                    WindowLog.Ecrire("   Type : " + f.GetTypeName2());
                }

            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }
    }
}
