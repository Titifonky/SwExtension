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
                Face2 Face = MdlBase.eSelect_RecupererObjet<Face2>();
                if (Face.IsNull()) return;

                Body2 CorpsBase = Face.GetBody();
                if (CorpsBase.IsNull()) return;

                Component2 cpCorpsBase = MdlBase.eSelect_RecupererComposant();
                String MateriauxCorpsBase = "";

                if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                {
                    cpCorpsBase = MdlBase.eComposantRacine();
                    MateriauxCorpsBase = CorpsBase.eGetMateriauCorpsOuPiece(MdlBase.ePartDoc(), MdlBase.eNomConfigActive());
                }
                else
                {
                    MateriauxCorpsBase = CorpsBase.eGetMateriauCorpsOuComp(cpCorpsBase);
                }

                MdlBase.eEffacerSelection();

                var ListeCorpsIdentiques = new List<Body2>();

                foreach (var comp in MdlBase.eComposantRacine().eRecListeComposant(c => { return c.TypeDoc() == eTypeDoc.Piece; }))
                {
                    foreach (var Corps in comp.eListeCorps())
                    {
                        var MateriauCorpsTest = Corps.eGetMateriauCorpsOuComp(comp);
                        if (MateriauxCorpsBase != MateriauCorpsTest) continue;

                        if (Corps.eComparerGeometrie(CorpsBase) == Sw.Comparaison_e.Semblable)
                        {
                            ListeCorpsIdentiques.Add(Corps);
                        }
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
