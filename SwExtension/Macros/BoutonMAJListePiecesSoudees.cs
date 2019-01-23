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
        ModuleTitre("MAJ Liste des pièces soudées"),
        ModuleNom("MAJListePiecesSoudees")]
    public class BoutonMAJListePiecesSoudees : BoutonBase
    {

        protected override void Command()
        {
            try
            {
                if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                {
                    PartDoc piece = MdlBase.ePartDoc();
                    piece.eInsererListeDesPiecesSoudees();
                    piece.eMajListeDesPiecesSoudees();
                    return;
                }


                if (MdlBase.TypeDoc() == eTypeDoc.Assemblage)
                    MdlBase.eRecParcourirComposants(Maj);
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        private Dictionary<String, String> _Dic = new Dictionary<string, string>();

        private Boolean Maj(Component2 Cp)
        {
            if (!Cp.IsHidden(true) && !_Dic.ContainsKey(Cp.eKeyAvecConfig()) && Cp.eEstDansLeDossier(MdlBase) && (Cp.TypeDoc() == eTypeDoc.Piece))
            {
                _Dic.Add(Cp.eKeyAvecConfig(), "");

                WindowLog.Ecrire(Cp.eNomAvecExt());
                try
                {
                    PartDoc piece = Cp.ePartDoc();
                    piece.eInsererListeDesPiecesSoudees();
                    piece.eMajListeDesPiecesSoudees();
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }

                
            }

            return false;
        }
    }
}
