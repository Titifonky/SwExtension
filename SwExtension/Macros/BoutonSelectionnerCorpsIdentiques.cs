using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Selectionner les corps identiques"),
        ModuleNom("SelectionnerCorpsIdentiques")]
    public class BoutonSelectionnerCorpsIdentiques : BoutonBase
    {
        public BoutonSelectionnerCorpsIdentiques()
        {
            LogToWindowLog = true;
        }

        protected override void Command()
        {
            try
            {
                ModelDoc2 mdl = App.ModelDoc2;

                Face2 Face = mdl.eSelect_RecupererObjet<Face2>();
                if (Face.IsNull()) return;

                Body2 CorpsBase = Face.GetBody();
                if (CorpsBase.IsNull()) return;

                String MateriauxCorpsBase = CorpsBase.eGetMateriauCorpsOuComp(mdl.eSelect_RecupererComposant());
                mdl.eEffacerSelection();

                var ListeCorpsIdentiques = new List<Body2>();

                foreach (var comp in mdl.eComposantRacine().eRecListeComposant(c => { return c.TypeDoc() == eTypeDoc.Piece; }))
                {
                    foreach (var Corps in comp.eListeCorps())
                    {
                        var MateriauCorpsTest = Corps.eGetMateriauCorpsOuComp(comp);
                        if (MateriauxCorpsBase != MateriauCorpsTest) continue;

                        if (Corps.eEstSemblable(CorpsBase))
                            ListeCorpsIdentiques.Add(Corps);
                    }
                }

                WindowLog.EcrireF("Nb de corps identiques : {0}", ListeCorpsIdentiques.Count);
                foreach (var corps in ListeCorpsIdentiques)
                {
                    corps.eSelect(true);
                }
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }
    }
}
