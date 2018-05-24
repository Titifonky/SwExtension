using LogDebugging;
using Outils;
using SwExtension;
using System;
using System.IO;

namespace ModuleLaser
{
    namespace ModuleNumeroterDossier
    {
        [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
            ModuleTitre("Numeroter les dossiers"),
            ModuleNom("NumeroterDossier"),
            ModuleDescription("Numeroter les dossiers.")
            ]
        public class PageNumeroterDossier : BoutonPMPManager
        {
            private Parametre CombinerCorpsIdentiques;

            public PageNumeroterDossier()
            {
                try
                {
                    CombinerCorpsIdentiques = _Config.AjouterParam("CombinerCorpsIdentiques", false, "Combiner les corps identiques des différents modèles");

                    OnCalque += Calque;
                    OnRunOkCommand += RunOkCommand;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private CtrlCheckBox _CheckBox_CombinerCorps;

            protected void Calque()
            {
                try
                {
                    Groupe G;

                    G = _Calque.AjouterGroupe("Options");

                    _CheckBox_CombinerCorps = G.AjouterCheckBox(CombinerCorpsIdentiques);
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            protected void RunOkCommand()
            {
                CmdNumeroterDossier Cmd = new CmdNumeroterDossier();

                Cmd.MdlBase = App.Sw.ActiveDoc;
                Cmd.CombinerCorpsIdentiques = _CheckBox_CombinerCorps.IsChecked;

                Cmd.Executer();
            }
        }
    }
}
