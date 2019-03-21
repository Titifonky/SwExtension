using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Masquer les corps de la vue"),
        ModuleNom("MasquerCorpsVue")]
    public class BoutonMasquerCorpsVue : BoutonBase
    {

        protected override void Command()
        {
            try
            {
                View Vue = MdlBase.eSelect_RecupererObjet<View>(1, -1);

                if (Vue.GetBodiesCount() == 0)
                {
                    WindowLog.Ecrire("Aucun corps");
                    return;
                }

                Object[] TabCorps = (Object[])Vue.Bodies;

                DispatchWrapper[] arrBodiesIn = new DispatchWrapper[Vue.GetBodiesCount()];

                for (int i = 0; i < TabCorps.Length; i++)
                {
                    arrBodiesIn[i] = new DispatchWrapper(TabCorps[i]);
                    WindowLog.Ecrire(((Body2)TabCorps[i]).Name);
                }

                Vue.Bodies = (arrBodiesIn);

                WindowLog.Ecrire(Vue.Name);
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }
    }
}
