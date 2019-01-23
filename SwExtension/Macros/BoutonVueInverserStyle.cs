using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Inverser le style d'affichage"),
        ModuleNom("VueInverserStyle")]
    public class BoutonVueInverserStyle : BoutonBase
    {
        protected override void Command()
        {
            try
            {
                var ListeVue = MdlBase.eSelect_RecupererListeObjets<View>();

                if (ListeVue.IsNull() || (ListeVue.Count == 0))
                    ListeVue = MdlBase.eDrawingDoc().eFeuilleActive().eListeDesVues();

                if (ListeVue.IsNull() || (ListeVue.Count == 0)) return;

                MdlBase.eEffacerSelection();

                foreach (var vue in ListeVue)
                {
                    var style = (swDisplayMode_e)vue.GetDisplayMode2();
                    switch (style)
                    {
                        case swDisplayMode_e.swHIDDEN:
                            vue.SetDisplayMode3(false, (int)swDisplayMode_e.swSHADED, false, true);
                            break;
                        case swDisplayMode_e.swSHADED:
                            vue.SetDisplayMode3(false, (int)swDisplayMode_e.swHIDDEN, false, true);
                            break;
                        default:
                            break;
                    }
                }

            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }
    }
}
