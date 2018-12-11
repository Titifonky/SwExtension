using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.IO;

namespace ModuleProduction
{
    namespace ModuleRepererDossier
    {
        [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
            ModuleTitre("Repérer les dossiers"),
            ModuleNom("RepererDossier"),
            ModuleDescription("Repérer les dossiers.")
            ]
        public class PageRepererDossier : BoutonPMPManager
        {
            private Parametre CombinerCorpsIdentiques;

            private ModelDoc2 MdlBase = null;
            private int IndiceCampagne = 1;
            private String DossierPiece = "";
            private String FichierNomenclature = "";

            public PageRepererDossier()
            {
                try
                {
                    CombinerCorpsIdentiques = _Config.AjouterParam("CombinerCorpsIdentiques", false, "Combiner les corps identiques des différents modèles");

                    MdlBase = App.Sw.ActiveDoc;
                    OnCalque += Calque;
                    OnRunAfterActivation += Rechercher_Infos;
                    OnRunOkCommand += RunOkCommand;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private CtrlCheckBox _CheckBox_CombinerCorps;
            private CtrlCheckBox _CheckBox_MajDossiers;

            protected void Calque()
            {
                try
                {
                    Groupe G;

                    G = _Calque.AjouterGroupe("Options");

                    _CheckBox_CombinerCorps = G.AjouterCheckBox(CombinerCorpsIdentiques);
                    _CheckBox_MajDossiers = G.AjouterCheckBox("Supprimer les repères existants");
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            protected void Rechercher_Infos()
            {
                WindowLog.Ecrire("Recherche des éléments existants :");

                // Recherche du dernier indice de la campagne de repérage

                // Création du dossier pièces s'il n'existe pas
                DossierPiece = MdlBase.CreerDossier(OutilsCommun.DossierPieces);
                // Recherche de la nomenclature
                FichierNomenclature = MdlBase.CreerFichierTexte(OutilsCommun.DossierPieces, OutilsCommun.FichierNomenclature);
            }

            protected void RunOkCommand()
            {
                CmdRepererDossier Cmd = new CmdRepererDossier();

                Cmd.MdlBase = App.Sw.ActiveDoc;
                Cmd.CombinerCorpsIdentiques = _CheckBox_CombinerCorps.IsChecked;
                Cmd.MajDossiers = _CheckBox_MajDossiers.IsChecked;

                Cmd.Executer();
            }
        }
    }
}
