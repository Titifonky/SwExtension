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
            DrawingDoc dessin = App.ModelDoc2.eDrawingDoc();

            String nom = dessin.eFeuilleActive().GetName();

            if (Interaction.InputBox("Nouveau nom", "Nom :", ref nom) == DialogResult.OK)
            {
                if (!String.IsNullOrWhiteSpace(nom))
                    dessin.eFeuilleActive().SetName(nom);
            }
        }
    }
}
