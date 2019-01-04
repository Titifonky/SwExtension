using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Retourner les Dvps selectionnés"),
        ModuleNom("RetournerDvp")]
    public class BoutonRetournerDvp : BoutonBase
    {
        private ModelDoc2 MdlBase = null;

        protected override void Command()
        {
            try
            {
                MdlBase = App.ModelDoc2;

                var ListeVue = MdlBase.eSelect_RecupererListeObjets<View>();

                if (ListeVue.IsNull() || (ListeVue.Count == 0))
                    ListeVue = MdlBase.eDrawingDoc().eFeuilleActive().eListeDesVues();

                if (ListeVue.IsNull() || (ListeVue.Count == 0)) return;

                MdlBase.eEffacerSelection();

                foreach (var vue in ListeVue)
                {
                    if (vue.IsNull()) continue;

                    WindowLog.Ecrire(vue.Name);
                    vue.FlipView = !vue.FlipView;
                }

            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }
    }
}
