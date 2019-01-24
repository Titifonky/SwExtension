using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Modifier les lumières"),
        ModuleNom("ModifierLumiere")]
    public class BoutonModifierLumiere : BoutonBase
    {
        public BoutonModifierLumiere()
        {
            LogToWindowLog = false;
        }

        protected override void Command()
        {
            try
            {
                MdlBase.SetLightSourcePropertyValuesVB("Ambiante-1", 1, 0, 16777215, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.25, 0, 0, false);
                MdlBase.LockLightToModel(0, true);
                MdlBase.SetLightSourcePropertyValuesVB("Directionnelle-1", 4, 0.48, 16777215, 1, -0.34344602761096943, 0.3759484815317567, 0.77139240786555829, 0, 0, 0, 0, 0, 0, 0, 0.3, 0.3, 0, false);
                MdlBase.LockLightToModel(1, true);
                MdlBase.SetLightSourcePropertyValuesVB("Directionnelle-1", 4, 0.05, 16777215, 1, 0.88322349334397254, -0.22360948011027768, -0.15573613187212906, 0, 0, 0, 0, 0, 0, 0, 0.05, 0.05, 0, false);
                MdlBase.LockLightToModel(2, true);
                MdlBase.SetLightSourcePropertyValuesVB("Directionnelle-1", 4, 0.2, 16777215, 1, 0.6313116669339, -0.2392275902706372, 0.63131166693389773, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, false);
                MdlBase.LockLightToModel(3, true);
                MdlBase.GraphicsRedraw();
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }
}
