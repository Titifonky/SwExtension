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
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Contrôler le reperage"),
        ModuleNom("ControlerReperage")]

    public class BoutonControlerReperage : BoutonBase
    {
        protected override void Command()
        {
            try
            {
                ModelDoc2 MdlBase = App.ModelDoc2;

                var ListeCorpsExistant = MdlBase.pChargerNomenclature();

                if (ListeCorpsExistant.Count > 0)
                {
                    var IndiceCampagne = ListeCorpsExistant.First().Value.Campagne.Keys.Max();
                    var aff = new AffichageElementWPF(ListeCorpsExistant, IndiceCampagne);
                    aff.ShowDialog();
                }

            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }
    }
}
