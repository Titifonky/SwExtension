using LogDebugging;
using Outils;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace ModuleLierLesConfigurations
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Créer/Lier les configurations"),
        ModuleNom("LierLesConfigurations"),
        ModuleDescription("Lier les configurations des composants selectionnés au composant racine")
        ]
    public class PageLierLesConfigurations : BoutonPMPManager
    {
        private Parametre _pCreerConfig;
        private Parametre _pListeConfigs;
        private Parametre _pSupprimerNvlFonction;
        private Parametre _pSupprimerNvComposant;

        public PageLierLesConfigurations()
        {
            _pCreerConfig = _Config.AjouterParam("CreerConfig", true, "Creer les configs manquante");
            _pListeConfigs = _Config.AjouterParam("ListeConfigs", "0 1 2 3 4 5", "Ajouter des configs" + "\r\n Noms séparés par un espace");
            _pSupprimerNvlFonction = _Config.AjouterParam("SupprimerNvlFonction", true, "Supprimer les nouvelles fonctions ou contraintes\r\n dans la configuration");
            _pSupprimerNvComposant = _Config.AjouterParam("SupprimerNvComposant", false, "Supprimer les nouveaux composants\r\n dans la configuration");

            OnCalque += Calque;
            OnRunOkCommand += RunOkCommand;
        }

        private CtrlTextBox _Texte_ListeConfigs;
        private CtrlSelectionBox _Select_Composants;
        private CtrlCheckBox _CheckBox_CreerLesConfigsManquantes;
        private CtrlCheckBox _CheckBox_SupprimerNvlFonction;
        private CtrlCheckBox _CheckBox_SupprimerNvComposant;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Ajouter des configs : ");

                _Texte_ListeConfigs = G.AjouterTexteBox(_pListeConfigs);

                if (MdlBase.TypeDoc() != eTypeDoc.Piece)
                {
                    G = _Calque.AjouterGroupe("Selectionner les composants à lier");

                    _Select_Composants = G.AjouterSelectionBox("Selectionnez les composants");
                    _Select_Composants.SelectionMultipleMemeEntite = false;
                    _Select_Composants.SelectionDansMultipleBox = false;
                    _Select_Composants.UneSeuleEntite = false;
                    _Select_Composants.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
                    // Filtre sur les composants autre que root
                    //_Select_Composants.OnSubmitSelection += SelectionnerComposantsParent;
                    _Select_Composants.Hauteur = 8;
                }

                G = _Calque.AjouterGroupe("Options");

                if (MdlBase.TypeDoc() != eTypeDoc.Piece)
                    _CheckBox_CreerLesConfigsManquantes = G.AjouterCheckBox(_pCreerConfig);

                _CheckBox_SupprimerNvlFonction = G.AjouterCheckBox(_pSupprimerNvlFonction);
                _CheckBox_SupprimerNvComposant = G.AjouterCheckBox(_pSupprimerNvComposant);
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdLierLesConfigurations Cmd = new CmdLierLesConfigurations();
            Cmd.MdlBase = MdlBase;
            
            if (MdlBase.TypeDoc() != eTypeDoc.Piece)
                Cmd.ListeComposants = MdlBase.eSelect_RecupererListeComposants(_Select_Composants.Marque);


            Cmd.CreerConfigsManquantes = true;

            if (MdlBase.TypeDoc() != eTypeDoc.Piece)
                Cmd.CreerConfigsManquantes = _CheckBox_CreerLesConfigsManquantes.IsChecked;

            Cmd.SupprimerNvlFonction = _CheckBox_SupprimerNvlFonction.IsChecked;
            Cmd.SupprimerNvComposant = _CheckBox_SupprimerNvComposant.IsChecked;

            Cmd.ListeConfig = new List<string>(_Texte_ListeConfigs.Text.Trim().Split(' '));

            MdlBase.ClearSelection2(true);

            Cmd.Executer();
        }
    }
}
