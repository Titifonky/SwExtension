using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Windows.Forms;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Renommer la feuille"),
        ModuleNom("RenommerFeuille")]
    public class BoutonRenommerFeuille : BoutonBase
    {
        protected override void Command()
        {
            try
            {
                DrawingDoc dessin = MdlBase.eDrawingDoc();

                String nom = dessin.eFeuilleActive().GetName();

                if (Interaction.InputBox("Nouveau nom", "Nom :", ref nom) == DialogResult.OK)
                {
                    if (!String.IsNullOrWhiteSpace(nom))
                        dessin.eFeuilleActive().SetName(nom);
                }
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }
}
