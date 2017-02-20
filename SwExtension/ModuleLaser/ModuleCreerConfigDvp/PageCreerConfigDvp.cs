using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

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
        private Parametre NumeroterDossier;
        private Parametre ToutesLesConfigurations;
        private Parametre MasquerEsquisses;
        private Parametre SupprimerFonctions;
        private Parametre NomFonctionSupprimer;

        public PageCreerConfigDvp()
        {
            SupprimerLesAnciennesConfigs = _Config.AjouterParam("SupprimerLesAnciennesConfigs", false, "Supprimer les anciennes configs dvp");
            ReconstuireLesConfigs = _Config.AjouterParam("ReconstuireLesConfig", false, "Reconstruire toutes les configs dvp");
            NumeroterDossier = _Config.AjouterParam("NumeroterDossier", true, "Numeroter les dossier");
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
        private CtrlCheckBox _CheckBox_NumeroterDossier;
        private CtrlCheckBox _CheckBox_ReinitialiserNoDossier;
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
                _CheckBox_NumeroterDossier = G.AjouterCheckBox(NumeroterDossier);
                _CheckBox_ReinitialiserNoDossier = G.AjouterCheckBox("Reinitialiser les n° de dossier");
                _CheckBox_NumeroterDossier.OnUnCheck += _CheckBox_ReinitialiserNoDossier.UnCheck;
                _CheckBox_NumeroterDossier.OnIsCheck += _CheckBox_ReinitialiserNoDossier.IsEnable;
                _CheckBox_ReinitialiserNoDossier.IsEnabled = _CheckBox_NumeroterDossier.IsChecked;
                _CheckBox_ReinitialiserNoDossier.OnCheck += _CheckBox_ToutesLesConfigurations.Check;
                _CheckBox_ReinitialiserNoDossier.StdIndent();

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
            NumeroterDossier.SetValeur<Boolean>(false);

            CmdCreerConfigDvp Cmd = new CmdCreerConfigDvp();
            Cmd.MdlBase = App.ModelDoc2;

            Cmd.SupprimerLesAnciennesConfigs = _CheckBox_SupprimerLesAnciennesConfigs.IsChecked;
            Cmd.ReconstuireLesConfigs = _CheckBox_ReconstuireLesConfigs.IsChecked;
            Cmd.MasquerEsquisses = _CheckBox_MasquerEsquisses.IsChecked;
            Cmd.NumeroterDossier = _CheckBox_NumeroterDossier.IsChecked;
            Cmd.ReinitialiserNoDossier = _CheckBox_NumeroterDossier.IsEnabled ? _CheckBox_ReinitialiserNoDossier.IsChecked : false;
            Cmd.ToutesLesConfigurations = _CheckBox_ToutesLesConfigurations.IsChecked;
            Cmd.SupprimerFonctions = _GroupeAvecCheckBox.IsChecked;
            Cmd.NomFonctionSupprimer = _TextBox_NomFonctionSupprimer.Text;

            Cmd.Executer();
        }

    }
}
