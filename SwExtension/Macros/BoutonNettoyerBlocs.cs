using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Nettoyer les blocs"),
        ModuleNom("NettoyerBlocs")]

    public class BoutonNettoyerBlocs : BoutonBase
    {
        protected override void Command()
        {
            try
            {
                WindowLog.Ecrire("Nettoyer les blocs :");

                SupprimerDefBloc(MdlBase);

                WindowLog.SautDeLigne();
                WindowLog.Ecrire("Nettoyage terminé");
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }

        private void SupprimerDefBloc(ModelDoc2 mdl)
        {
            var TabDef = (Object[])mdl.SketchManager.GetSketchBlockDefinitions();
            if (TabDef.IsRef())
            {
                foreach (SketchBlockDefinition blocdef in TabDef)
                {
                    Feature f = blocdef.GetFeature();
                    WindowLog.EcrireF(" - {0}", f.Name);

                    if (blocdef.GetInstanceCount() > 0) continue;

                    f.eSelect();
                    var res = mdl.Extension.DeleteSelection2((int)(swDeleteSelectionOptions_e.swDelete_Absorbed | swDeleteSelectionOptions_e.swDelete_Children));
                    WindowLog.EcrireF("    Supprimée : {0}", res);
                    mdl.eEffacerSelection();
                }
            }
        }
    }
}
