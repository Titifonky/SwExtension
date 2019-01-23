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
            LogToWindowLog = false;
        }

        protected override void Command()
        {
            try
            {
                String NomProp = "Date";

                CustomPropertyManager PM = MdlBase.eGestProp("");

                if(MdlBase.ePropExiste(NomProp))
                    PM.Add3(NomProp, (int)swCustomInfoType_e.swCustomInfoDate, DateTime.Today.ToString("d"), (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                else
                    PM.Add3(NomProp, (int)swCustomInfoType_e.swCustomInfoDate, DateTime.Today.ToString("d"), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);

                MdlBase.EditRebuild3();
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }
}
