using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ModuleProduction.ModuleGenererConfigDvp
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Générer les config. dvp"),
        ModuleNom("GenererConfigDvp"),
        ModuleDescription("Générer les configurations dvp de chaque tole"),
        ModuleAide("")
        ]
    public class PageGenererConfigDvp : BoutonPMPManager
    {
        private Parametre SupprimerLesAnciennesConfigs;

        public PageGenererConfigDvp()
        {
            SupprimerLesAnciennesConfigs = _Config.AjouterParam("SupprimerLesAnciennesConfigs", false, "Supprimer les anciennes configs dvp");

            OnCalque += Calque;
            OnRunOkCommand += RunOkCommand;
        }

        private CtrlCheckBox _CheckBox_SupprimerLesAnciennesConfigs;

        protected void Calque()
        {
            try
            {
                Groupe G;
                G = _Calque.AjouterGroupe("Options");

                _CheckBox_SupprimerLesAnciennesConfigs = G.AjouterCheckBox(SupprimerLesAnciennesConfigs);

            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdGenererConfigDvp Cmd = new CmdGenererConfigDvp();

            Cmd.MdlBase = App.ModelDoc2;
            Cmd.SupprimerLesAnciennesConfigs = _CheckBox_SupprimerLesAnciennesConfigs.IsChecked;

            Cmd.Executer();
        }

    }
}
