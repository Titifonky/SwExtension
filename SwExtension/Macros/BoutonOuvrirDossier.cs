﻿using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Ouvrir le dossier du composant"),
        ModuleNom("OuvrirDossier")]
    public class BoutonOuvrirDossier : BoutonBase
    {
        public BoutonOuvrirDossier()
        {
            LogToWindowLog = false;
        }

        protected override void Command()
        {
            try
            {
                System.Diagnostics.Process.Start(MdlBase.eDossier());
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }
}
