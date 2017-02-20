﻿using Outils;
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

            String nom = "";

            if (Interaction.InputBox("Nouveau nom", "Nom :", ref nom) == DialogResult.OK)
            {
                if (!String.IsNullOrWhiteSpace(nom))
                {
                    Sheet Feuille = (Sheet)dessin.GetCurrentSheet();
                    Feuille.SetName(nom);
                }
            }
        }
    }
}
