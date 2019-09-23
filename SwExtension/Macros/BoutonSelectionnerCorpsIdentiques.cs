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
                if (Face.IsNull())
                {
                    WindowLog.Ecrire("Pas de face selectionnée");
                    return;
                }

                Body2 CorpsBase = Face.GetBody();
                if (CorpsBase.IsNull())
                {
                    WindowLog.Ecrire("Pas de corps selectionnée");
                    return;
                }

                String MateriauxCorpsBase = "";

                if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                    MateriauxCorpsBase = CorpsBase.eGetMateriauCorpsOuPiece(MdlBase.ePartDoc(), MdlBase.eNomConfigActive());
                else
                {
                    Component2 cpCorpsBase = MdlBase.eSelect_RecupererComposant();
                    MateriauxCorpsBase = CorpsBase.eGetMateriauCorpsOuComp(cpCorpsBase);
                }

                MdlBase.eEffacerSelection();

                WindowLog.Ecrire("Matériau : " + MateriauxCorpsBase);

                var ListeCorpsIdentiques = new List<Body2>();

                Action<Body2, String> Test = delegate (Body2 corps, String mat)
                {
                    if (MateriauxCorpsBase != mat) return;

                    if (corps.eComparerGeometrie(CorpsBase) == Sw.Comparaison_e.Semblable)
                        ListeCorpsIdentiques.Add(corps);
                };

                if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                {
                    var Piece = MdlBase.ePartDoc();
                    foreach (var Corps in Piece.eListeCorps())
                    {
                        var MateriauCorpsTest = Corps.eGetMateriauCorpsOuPiece(Piece, MdlBase.eNomConfigActive());
                        Test(Corps, MateriauCorpsTest);
                    }
                }
                else
                {
                    foreach (var comp in MdlBase.eComposantRacine().eRecListeComposant(c => { return c.TypeDoc() == eTypeDoc.Piece; }))
                    {
                        foreach (var Corps in comp.eListeCorps())
                        {
                            var MateriauCorpsTest = Corps.eGetMateriauCorpsOuComp(comp);
                            Test(Corps, MateriauCorpsTest);
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
