using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Supp. Config depliee"),
        ModuleNom("SupprimerConfigDepliee")]
    public class BoutonSupprimerConfigDepliee : BoutonBase
    {
        private ModelDoc2 MdlBase = null;

        protected override void Command()
        {
            MdlBase = App.ModelDoc2;

            try
            {
                if (App.ModelDoc2.TypeDoc() == eTypeDoc.Piece)
                {
                    SupprimerConfigs(App.ModelDoc2.eComposantRacine());
                    return;
                }


                if (App.ModelDoc2.TypeDoc() == eTypeDoc.Assemblage)
                    App.ModelDoc2.eRecParcourirComposants(SupprimerConfigs);
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private Dictionary<String, String> _Dic = new Dictionary<string, string>();

        private Boolean SupprimerConfigs(Component2 Cp)
        {
            try
            {
                if ((Cp.TypeDoc() == eTypeDoc.Piece) && !Cp.IsHidden(true) && !_Dic.ContainsKey(Cp.eKeySansConfig()) && Cp.eEstDansLeDossier(MdlBase))
                {
                    _Dic.Add(Cp.eKeySansConfig(), "");

                    WindowLog.Ecrire(Cp.eNomAvecExt());
                    foreach (Configuration Cf in Cp.eModelDoc2().eListeConfigs(eTypeConfig.Depliee))
                    {
                        String IsSup = Cf.eSupprimerConfigAvecEtatAff(Cp.eModelDoc2()) ? "Ok" : "Erreur";
                        WindowLog.EcrireF("  {0} : {1}", Cf.Name, IsSup);
                    }
                }
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }

            return false;
        }
    }
}
