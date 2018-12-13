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
            private int _IndiceCampagne = 1;
            private int IndiceMin = 0;
            private String DossierPiece = "";
            private String FichierNomenclature = "";

            private int IndiceCampagne
            {
                get { return _IndiceCampagne; }
                set
                {
                    _IndiceCampagne = value;
                    _Texte_IndiceCampagne.Text = _IndiceCampagne.ToString();
                }
            }

            public PageRepererDossier()
            {
                try
                {
                    CombinerCorpsIdentiques = _Config.AjouterParam("CombinerCorpsIdentiques", false, "Combiner les corps identiques des différents modèles");

                    MdlBase = App.Sw.ActiveDoc;
                    OnCalque += Calque;
                    OnRunAfterActivation += RechercherIndiceCampagne;
                    OnRunOkCommand += RunOkCommand;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private CtrlTextBox _Texte_IndiceCampagne;
            private CtrlCheckBox _CheckBox_CombinerCorps;
            private CtrlCheckBox _CheckBox_SupprimerReperes;

            protected void Calque()
            {
                try
                {
                    Groupe G;

                    G = _Calque.AjouterGroupe("Options");

                    _Texte_IndiceCampagne = G.AjouterTexteBox("Indice de la campagne de repérage :");
                    _Texte_IndiceCampagne.LectureSeule = true;
                    _CheckBox_CombinerCorps = G.AjouterCheckBox(CombinerCorpsIdentiques);
                    _CheckBox_SupprimerReperes = G.AjouterCheckBox("Supprimer les repères de la précédente campagne");
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            protected void RechercherIndiceCampagne()
            {
                IndiceCampagne = 1;

                WindowLog.Ecrire("Recherche des éléments existants :");

                // Recherche du dernier indice de la campagne de repérage

                // Création du dossier pièces s'il n'existe pas
                MdlBase.CreerDossier(OutilsCommun.DossierPieces, out DossierPiece);
                // Recherche de la nomenclature
                MdlBase.CreerFichierTexte(OutilsCommun.DossierPieces, OutilsCommun.FichierNomenclature, out FichierNomenclature);

                using (var sr = new StreamReader(FichierNomenclature))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        
                    }
                }

                if(IndiceMin == 0)
                {
                    _CheckBox_SupprimerReperes.IsChecked = true;
                    _CheckBox_SupprimerReperes.IsEnabled = false;
                }
            }

            protected void RunOkCommand()
            {
                CmdRepererDossier Cmd = new CmdRepererDossier();

                Cmd.MdlBase = App.Sw.ActiveDoc;
                Cmd.IndiceCampagne = IndiceCampagne;
                Cmd.Indice = IndiceMin;
                Cmd.CombinerCorpsIdentiques = _CheckBox_CombinerCorps.IsChecked;
                Cmd.SupprimerReperes = _CheckBox_SupprimerReperes.IsChecked;

                Cmd.Executer();
            }
        }
    }
}
