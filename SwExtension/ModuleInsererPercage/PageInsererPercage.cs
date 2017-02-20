using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ModuleInsererPercage
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage),
        ModuleTitre("Inserer les perçages"),
        ModuleNom("InsererPercage"),
        ModuleDescription("Inserer les perçages")
        ]
    public class PageInsererPercage : BoutonPMPManager
    {
        private Parametre _pPieceBase;
        private Parametre _pPercage;
        private Parametre _pDiametre;
        private Parametre _pPercageOuvert;
        private Parametre _pToutesLesConfig;

        public PageInsererPercage()
        {
            _pPieceBase = _Config.AjouterParam("PieceBase", "PP01", "Selectionnez le composant de reference");
            _pPercage = _Config.AjouterParam("Percage", "PX50", "Selectionnez le composant de perçage");
            _pDiametre = _Config.AjouterParam("Diametre", "0", "Diametres des trous :");
            _pPercageOuvert = _Config.AjouterParam("PercageOuvert", false, "Prendre en compte les perçages ouvert");

            _pToutesLesConfig = _Config.AjouterParam("ToutesLesConfig", false, "Appliquer à toutes les configs");

            OnCalque += Calque;
            OnRunOkCommand += RunOkCommand;
        }

        private CtrlSelectionBox _Select_Base;
        private CtrlSelectionBox _Select_Percage;
        private CtrlSelectionBox _Select_Entite_Contrainte;
        private CtrlTextBox _Text_Diametre;
        private CtrlCheckBox _Check_PercageOuvert;

        private CtrlCheckBox _CheckBox_ToutesLesConfig;
        private CtrlCheckBox _CheckBox_EnregistrerSelection;
        private CtrlButton _Button_Preselection;

        protected void Calque()
        {
            try
            {
                Groupe G;
                G = _Calque.AjouterGroupe("Appliquer");

                _CheckBox_ToutesLesConfig = G.AjouterCheckBox(_pToutesLesConfig);
                _CheckBox_EnregistrerSelection = G.AjouterCheckBox("Enregistrer les selections");
                _Button_Preselection = G.AjouterBouton("Preselectionner");
                _Button_Preselection.OnButtonPress += delegate (object sender) { PreSelection(); };

                G = _Calque.AjouterGroupe("Piece sur laquelle inserer les perçages");

                _Select_Base = G.AjouterSelectionBox("Selectionnez le composant");
                _Select_Base.SelectionMultipleMemeEntite = false;
                _Select_Base.SelectionDansMultipleBox = false;
                _Select_Base.UneSeuleEntite = true;
                _Select_Base.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
                _Select_Base.OnSubmitSelection += SelectionnerPiece;
                _Select_Base.OnApplyOnSelection += delegate(Object Box) { SvgNomComposant(Box, _pPieceBase); };

                G = _Calque.AjouterGroupe("Composant de perçage");

                _Select_Percage = G.AjouterSelectionBox("Selectionnez le composant");
                _Select_Percage.SelectionMultipleMemeEntite = false;
                _Select_Percage.SelectionDansMultipleBox = false;
                _Select_Percage.UneSeuleEntite = true;
                _Select_Percage.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
                _Select_Percage.OnSubmitSelection += SelectionnerComposant1erNvx;
                _Select_Percage.OnApplyOnSelection += delegate(Object Box) { SvgNomComposant(Box, _pPercage); };

                _Select_Base.OnApplyOnSelection += _Select_Percage.GainedFocus;

                G = _Calque.AjouterGroupe("Plan ou face pour la contrainte de base"
                                            + "\r\n  Aucun(e) pour contraindre sur une face adjacente au trou"
                                            + "\r\n  Une face pour inserer seulement sur les trous de celle ci");

                _Select_Entite_Contrainte = G.AjouterSelectionBox("Selectionnez le plan ou la face");
                _Select_Entite_Contrainte.SelectionMultipleMemeEntite = false;
                _Select_Entite_Contrainte.SelectionDansMultipleBox = false;
                _Select_Entite_Contrainte.UneSeuleEntite = false;
                _Select_Entite_Contrainte.FiltreSelection(swSelectType_e.swSelDATUMPLANES, swSelectType_e.swSelFACES);

                _Select_Percage.OnApplyOnSelection += _Select_Entite_Contrainte.GainedFocus;

                G = _Calque.AjouterGroupe("Diametres des trous à contraindre en mm"
                                           + "\r\n  0 ou vide pour tout les perçages"
                                           + "\r\n  Valeurs séparés par une virgule");

                _Text_Diametre = G.AjouterTexteBox(_pDiametre, false);

                _Select_Entite_Contrainte.OnApplyOnSelection += _Text_Diametre.GainedFocus;

                G = _Calque.AjouterGroupe("Supprimer le perçage de base");

                _Check_PercageOuvert = G.AjouterCheckBox(_pPercageOuvert);

                // OnCheck, on enregistre les parametres
                _CheckBox_EnregistrerSelection.OnCheck += _Select_Base.ApplyOnSelection;
                _CheckBox_EnregistrerSelection.OnCheck += _Select_Percage.ApplyOnSelection;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        public void SvgNomComposant(Object SelBox, Parametre Param)
        {
            Component2 Cp = App.ModelDoc2.eSelect_RecupererComposant(1, ((CtrlSelectionBox)SelBox).Marque);
            if (Cp.IsRef() && _CheckBox_EnregistrerSelection.IsChecked)
                Param.SetValeur(Cp.eNomSansExt());
        }

        public void SvgNomFonction(Object SelBox, Parametre Param)
        {
            Feature F = App.ModelDoc2.eSelect_RecupererObjet<Feature>(1, ((CtrlSelectionBox)SelBox).Marque);
            if (F.IsRef() && _CheckBox_EnregistrerSelection.IsChecked)
                Param.SetValeur(F.Name);
        }

        protected void PreSelection()
        {
            try
            {
                App.ModelDoc2.ClearSelection2(true);

                SelectionMgr SelMgr = App.ModelDoc2.SelectionManager;
                Component2 Piece = App.ModelDoc2.eRecChercherComposant(c => { return Regex.IsMatch(c.Name2, _pPieceBase.GetValeur<String>())
                                                                            && !c.IsSuppressed(); });
                Component2 Percage = App.ModelDoc2.eRecChercherComposant(c => { return Regex.IsMatch(c.Name2, _pPercage.GetValeur<String>())
                                                                            && !c.IsSuppressed(); });

                if (Piece.IsRef())
                    App.ModelDoc2.eSelectMulti(Piece, _Select_Percage.Marque, false);

                if (Percage.IsRef())
                    App.ModelDoc2.eSelectMulti(Percage, _Select_Percage.Marque, false);

                _Select_Base.Focus = true;

            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdInsererPercage Cmd = new CmdInsererPercage();
            Cmd.MdlBase = App.ModelDoc2;
            Cmd.Base = App.ModelDoc2.eSelect_RecupererComposant(1, _Select_Base.Marque);
            Cmd.Percage = App.ModelDoc2.eSelect_RecupererComposant(1, _Select_Percage.Marque);
            Cmd.Face = App.ModelDoc2.eSelect_RecupererObjet<Face2>(1, _Select_Entite_Contrainte.Marque);
            Cmd.Plan = App.ModelDoc2.eSelect_RecupererObjet<Feature>(1, _Select_Entite_Contrainte.Marque);
            Cmd.ListeDiametre = new List<double>(_Text_Diametre.Text.Split(',').Select(x => { return x.Trim().eToDouble(); }));
            Cmd.PercageOuvert = _Check_PercageOuvert.IsChecked;
            Cmd.SurTouteLesConfigs = _CheckBox_ToutesLesConfig.IsChecked;

            App.ModelDoc2.ClearSelection2(true);

            Cmd.Executer();
        }

    }
}
