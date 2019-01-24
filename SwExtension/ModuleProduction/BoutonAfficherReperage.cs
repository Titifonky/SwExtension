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
        ModuleTitre("Afficher le reperage"),
        ModuleNom("AfficherReperage")]

    public class BoutonAfficherReperage : BoutonBase
    {
        public BoutonAfficherReperage()
        {
            LogToWindowLog = false;
        }

        protected override void Command()
        {
            try
            {
                var ListeCorpsExistant = MdlBase.pChargerNomenclature();

                if (ListeCorpsExistant.Count > 0)
                {
                    var IndiceCampagne = ListeCorpsExistant.First().Value.Campagne.Keys.Max();
                    var aff = new AffichageElementWPF(ListeCorpsExistant, IndiceCampagne);
                    aff.ShowDialog();
                }
                else
                {
                    WindowLog.Ecrire("Aucun repérage");
                }

            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }
    }
}
