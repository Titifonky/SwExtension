using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Date du jour"),
        ModuleNom("DateDuJour")]
    public class BoutonDateDuJour : BoutonBase
    {
        public BoutonDateDuJour()
        {
            LogToWindowLog = true;
        }

        protected override void Command()
        {
            try
            {
                MdlBase.ePropAdd("Date", DateTime.Today.ToString("d"));
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }
}
