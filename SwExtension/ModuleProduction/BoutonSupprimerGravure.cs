using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ModuleProduction
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Supprimer la gravure"),
        ModuleNom("SupprimerGravure")]

    public class BoutonSupprimerGravure : BoutonBase
    {
        protected override void Command()
        {
            try
            {
                MdlBase.eEffacerSelection();
                foreach (var vue in MdlBase.eDrawingDoc().eFeuilleActive().eListeDesVues())
                {
                    foreach (Annotation Ann in vue.GetAnnotations())
                    {
                        if (Ann.Layer == CONSTANTES.CALQUE_GRAVURE)
                            Ann.Select3(true, null);
                    }

                }

                MdlBase.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }
    }
}
