using LogDebugging;
using Outils;
using SwExtension;
using System;

namespace ModuleProduction.ModuleModifierDvp
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Modifier les dvps"),
        ModuleNom("ModifierDvp"),
        ModuleDescription("Modifier la présentation des dvps de tôle")
        ]
    public class PageModifierDvp : BoutonPMPManager
    {
        private Parametre AfficherLignePliage;
        private Parametre AfficherNotePliage;

        private Parametre InscrireNomTole;
        private Parametre TailleInscription;
        private Parametre OrienterDvp;
        private Parametre OrientationDvp;


        public PageModifierDvp()
        {
            try
            {
                AfficherLignePliage = _Config.AjouterParam("AfficherLignePliage", true, "Afficher les lignes de pliage");
                AfficherNotePliage = _Config.AjouterParam("AfficherNotePliage", true, "Afficher les notes de pliage");
                InscrireNomTole = _Config.AjouterParam("InscrireNomTole", true, "Inscrire la réf du dvp sur la tole");
                TailleInscription = _Config.AjouterParam("TailleInscription", 5, "Ht des inscriptions en mm", "Ht des inscriptions en mm");

                OrienterDvp = _Config.AjouterParam("OrienterDvp", false, "Orienter les dvps");
                OrientationDvp = _Config.AjouterParam("OrientationDvp", eOrientation.Portrait);

                OnCalque += Calque;
                OnRunOkCommand += RunOkCommand;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }
        
        private CtrlCheckBox _CheckBox_AfficherLignePliage;
        private CtrlCheckBox _CheckBox_AfficherNotePliage;
        private CtrlCheckBox _CheckBox_InscrireNomTole;
        private CtrlTextBox _Texte_TailleInscription;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Options");

                _CheckBox_AfficherLignePliage = G.AjouterCheckBox(AfficherLignePliage);
                _CheckBox_AfficherNotePliage = G.AjouterCheckBox(AfficherNotePliage);
                _CheckBox_AfficherNotePliage.StdIndent();

                _CheckBox_AfficherLignePliage.OnUnCheck += _CheckBox_AfficherNotePliage.UnCheck;
                _CheckBox_AfficherLignePliage.OnIsCheck += _CheckBox_AfficherNotePliage.IsEnable;

                // Pour eviter d'ecraser le parametre de "AfficherNotePliage", le met à jour seulement si
                if (!_CheckBox_AfficherLignePliage.IsChecked)
                {
                    _CheckBox_AfficherNotePliage.IsChecked = _CheckBox_AfficherLignePliage.IsChecked;
                    _CheckBox_AfficherNotePliage.IsEnabled = _CheckBox_AfficherLignePliage.IsChecked;
                }

                _CheckBox_InscrireNomTole = G.AjouterCheckBox(InscrireNomTole);
                _Texte_TailleInscription = G.AjouterTexteBox(TailleInscription, false);
                _Texte_TailleInscription.ValiderTexte += ValiderTextIsInteger;
                _Texte_TailleInscription.StdIndent();

                _CheckBox_InscrireNomTole.OnIsCheck += _Texte_TailleInscription.IsEnable;
                _Texte_TailleInscription.IsEnabled = _CheckBox_InscrireNomTole.IsChecked;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdModifierDvp Cmd = new CmdModifierDvp();

            Cmd.MdlBase = MdlBase;
            Cmd.AfficherLignePliage = _CheckBox_AfficherLignePliage.IsChecked;
            Cmd.AfficherNotePliage = _CheckBox_AfficherNotePliage.IsChecked;
            Cmd.InscrireNomTole = _CheckBox_InscrireNomTole.IsChecked;
            Cmd.TailleInscription = _Texte_TailleInscription.Text.eToInteger();

            Cmd.Executer();
        }
    }
}
