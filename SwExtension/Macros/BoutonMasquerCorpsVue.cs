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
        ModuleTitre("Masquer les corps de la vue"),
        ModuleNom("MasquerCorpsVue")]
    public class BoutonMasquerCorpsVue : BoutonBase
    {

        protected override void Command()
        {
            try
            {
                View Vue = MdlBase.eSelect_RecupererObjet<View>(1, -1);

                Body2[] TabCorps = (Body2[])Vue.Bodies;

                foreach (Body2 corps in TabCorps)
                {
                    WindowLog.Ecrire(corps.Name);
                }

                Vue.Bodies = Vue.Bodies;

                WindowLog.Ecrire(Vue.Name);
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }
    }
}
