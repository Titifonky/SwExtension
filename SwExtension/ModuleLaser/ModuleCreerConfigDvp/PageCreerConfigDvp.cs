using LogDebugging;
using Outils;
using SwExtension;
using System;

namespace ModuleLaser.ModuleCreerConfigDvp
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Creer les config. dvp"),
        ModuleNom("CreerConfigDvp"),
        ModuleDescription("Creer les configurations dvp de chaque tole"),
        ModuleAide("")
        ]
    public class PageCreerConfigDvp : BoutonPMPManager
    {
        private Parametre SupprimerLesAnciennesConfigs;
        private Parametre ReconstuireLesConfigs;
        private Parametre ToutesLesConfigurations;
        private Parametre MasquerEsquisses;
        private Parametre SupprimerFonctions;
        private Parametre NomFonctionSupprimer;

        public PageCreerConfigDvp()
        {
            SupprimerLesAnciennesConfigs = _Config.AjouterParam("SupprimerLesAnciennesConfigs", false, "Supprimer les anciennes configs dvp");
            ReconstuireLesConfigs = _Config.AjouterParam("ReconstuireLesConfig", false, "Reconstruire toutes les configs dvp");
            ToutesLesConfigurations = _Config.AjouterParam("ToutesLesConfigurations", false, "Appliquer à toutes les configs", "Creer les configs dvp pour toutes les configs pliées de chaque composants, même celles non utilisées dans le modele");
            MasquerEsquisses = _Config.AjouterParam("MasquerEsquisses", false, "Masquer toutes les esquisses");
            SupprimerFonctions = _Config.AjouterParam("SupprimerFonctions", false, "Supprimer les fonctions", "Supprimer les fonctions correspondant au motif donné");
            NomFonctionSupprimer = _Config.AjouterParam("NomFonctionSupprimer", "^S_", "Nom des fonctions à supprimer (Regex)");

            OnCalque += Calque;
            OnRunOkCommand += RunOkCommand;
        }

        private CtrlCheckBox _CheckBox_SupprimerLesAnciennesConfigs;
        private CtrlCheckBox _CheckBox_ReconstuireLesConfigs;
        private CtrlCheckBox _CheckBox_ToutesLesConfigurations;
        private CtrlCheckBox _CheckBox_MasquerEsquisses;
        private GroupeAvecCheckBox _GroupeAvecCheckBox;
        private CtrlTextBox _TextBox_NomFonctionSupprimer;

        protected void Calque()
        {
            try
            {
                Groupe G;
                G = _Calque.AjouterGroupe("Appliquer");

                _CheckBox_SupprimerLesAnciennesConfigs = G.AjouterCheckBox(SupprimerLesAnciennesConfigs);
                _CheckBox_ReconstuireLesConfigs = G.AjouterCheckBox(ReconstuireLesConfigs);

                _CheckBox_SupprimerLesAnciennesConfigs.OnCheck += _CheckBox_ReconstuireLesConfigs.UnCheck;
                _CheckBox_SupprimerLesAnciennesConfigs.OnIsCheck += _CheckBox_ReconstuireLesConfigs.IsDisable;

                // Pour eviter d'ecraser le parametre de "Reconstruire les configs", le met à jour seulement si
                if (_CheckBox_SupprimerLesAnciennesConfigs.IsChecked)
                {
                    _CheckBox_ReconstuireLesConfigs.IsChecked = !_CheckBox_SupprimerLesAnciennesConfigs.IsChecked;
                    _CheckBox_ReconstuireLesConfigs.IsEnabled = !_CheckBox_SupprimerLesAnciennesConfigs.IsChecked;
                }

                _CheckBox_ToutesLesConfigurations = G.AjouterCheckBox(ToutesLesConfigurations);

                G = _Calque.AjouterGroupe("Options");
                _CheckBox_MasquerEsquisses = G.AjouterCheckBox(MasquerEsquisses);

                _GroupeAvecCheckBox = _Calque.AjouterGroupeAvecCheckBox(SupprimerFonctions);
                _TextBox_NomFonctionSupprimer = _GroupeAvecCheckBox.AjouterTexteBox(NomFonctionSupprimer);
                _TextBox_NomFonctionSupprimer.IsEnabled = _GroupeAvecCheckBox.IsChecked;

                _GroupeAvecCheckBox.OnIsCheck += _TextBox_NomFonctionSupprimer.IsEnable;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdCreerConfigDvp Cmd = new CmdCreerConfigDvp();
            Cmd.MdlBase = MdlBase;

            Cmd.SupprimerLesAnciennesConfigs = _CheckBox_SupprimerLesAnciennesConfigs.IsChecked;
            Cmd.ReconstuireLesConfigs = _CheckBox_ReconstuireLesConfigs.IsChecked;
            Cmd.MasquerEsquisses = _CheckBox_MasquerEsquisses.IsChecked;
            Cmd.ToutesLesConfigurations = _CheckBox_ToutesLesConfigurations.IsChecked;
            Cmd.SupprimerFonctions = _GroupeAvecCheckBox.IsChecked;
            Cmd.NomFonctionSupprimer = _TextBox_NomFonctionSupprimer.Text;

            Cmd.Executer();
        }

    }
}
