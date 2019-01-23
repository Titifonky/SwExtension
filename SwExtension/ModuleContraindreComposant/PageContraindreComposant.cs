using LogDebugging;
using Outils;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace ModuleContraindreComposant
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage),
        ModuleTitre("Contraindre des composants"),
        ModuleNom("ContraindreComposant"),
        ModuleDescription("Contraindre les composants selectionnés au composant racine ou au composant de base sélectionné")
        ]
    public class PageContraindreComposant : BoutonPMPManager
    {
        private Parametre _pFixerComposant;

        public PageContraindreComposant()
        {
            _pFixerComposant = _Config.AjouterParam("FixerComposant", false, "Fixer les composants");

            OnCalque += Calque;
            OnRunOkCommand += RunOkCommand;
        }

        private CtrlSelectionBox _Select_CompBase;
        private CtrlSelectionBox _Select_Composants;
        private CtrlCheckBox _CheckBox_FixerComposant;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Selectionner le composant de base ou vide");

                _Select_CompBase = G.AjouterSelectionBox("Selectionnez le composant");
                _Select_CompBase.SelectionMultipleMemeEntite = false;
                _Select_CompBase.SelectionDansMultipleBox = false;
                _Select_CompBase.UneSeuleEntite = true;
                _Select_CompBase.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
                // Filtre sur les composants autre que root
                //_Select_CompBase.OnSubmitSelection += SelectionnerComposant1erNvx;

                G = _Calque.AjouterGroupe("Selectionner les composants à contraindre");

                _Select_Composants = G.AjouterSelectionBox("Selectionnez les composants");
                _Select_Composants.SelectionMultipleMemeEntite = false;
                _Select_Composants.SelectionDansMultipleBox = false;
                _Select_Composants.UneSeuleEntite = false;
                _Select_Composants.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
                // Filtre sur les composants autre que root
                _Select_Composants.OnSubmitSelection += SelectionnerComposant1erNvx;
                _Select_Composants.Hauteur = 8;

                _Select_Composants.Focus = true;

                G = _Calque.AjouterGroupe("Options");

                _CheckBox_FixerComposant = G.AjouterCheckBox(_pFixerComposant);

            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdContraindreComposant Cmd = new CmdContraindreComposant();
            Cmd.MdlBase = MdlBase;
            Cmd.CompBase = MdlBase.eSelect_RecupererComposant(1, _Select_CompBase.Marque);
            Cmd.ListeComposants = MdlBase.eSelect_RecupererListeComposants(_Select_Composants.Marque);
            Cmd.FixerComposant = _CheckBox_FixerComposant.IsChecked;

            MdlBase.ClearSelection2(true);

            Cmd.Executer();
        }
    }
}
