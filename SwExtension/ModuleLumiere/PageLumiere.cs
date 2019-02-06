using LogDebugging;
using Outils;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;

namespace ModuleLumiere
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Mettre à jour la lumière"),
        ModuleNom("Lumiere"),
        ModuleDescription("Mettre à jour la lumière du modèle")
        ]
    public class PageLumiere : BoutonPMPManager
    {
        private Parametre ValAmbiante;
        private Parametre ToutesLesConfigs;
        private Parametre DesactiverDirectionnelles;

        public PageLumiere()
        {
            try
            {
                ValAmbiante = _Config.AjouterParam("ValAmbiante", 0.85, "Valeur de la lumière ambiante");
                ToutesLesConfigs = _Config.AjouterParam("ToutesLesConfigs", true, "Appliquer à toutes les configs");
                DesactiverDirectionnelles = _Config.AjouterParam("DesactiverDirectionnelles", true, "Desactiver les directionnelles");

                OnCalque += Calque;
                OnRunOkCommand += RunOkCommand;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private CtrlTextBox _Texte_ValAmbiante;
        private CtrlCheckBox _CheckBox_ToutesLesConfigs;
        private CtrlCheckBox _CheckBox_DesactiverDirectionnelles;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Options");

                _Texte_ValAmbiante = G.AjouterTexteBox(ValAmbiante);
                _Texte_ValAmbiante.ValiderTexte += ValiderTextIsDouble;

                _CheckBox_ToutesLesConfigs = G.AjouterCheckBox(ToutesLesConfigs);
                _CheckBox_DesactiverDirectionnelles = G.AjouterCheckBox(DesactiverDirectionnelles);
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdLumiere Cmd = new CmdLumiere();

            Cmd.MdlBase = MdlBase;
            Cmd.ValAmbiante = _Texte_ValAmbiante.GetTextAs<Double>();
            Cmd.ToutesLesConfigs = _CheckBox_ToutesLesConfigs.IsChecked;
            Cmd.DesactiverDirectionnelles = _CheckBox_DesactiverDirectionnelles.IsChecked;

            Cmd.Executer();
        }
    }
}
