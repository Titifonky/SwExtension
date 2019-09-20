using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ModuleInsererPercageTole
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage),
        ModuleTitre("Inserer les perçages sur des toles"),
        ModuleNom("InsererPercageTole"),
        ModuleDescription("Inserer les perçages sur des toles")
        ]
    public class PageInsererPercageTole : BoutonPMPManager
    {
        private Parametre _pPieceBase;
        private Parametre _pPercage;
        private Parametre _pDiametre;
        private Parametre _pPercageOuvert;

        public PageInsererPercageTole()
        {
            _pPieceBase = _Config.AjouterParam("PieceBase", "PP01", "Selectionnez le composant de reference");
            _pPercage = _Config.AjouterParam("Percage", "PX50", "Selectionnez le composant de perçage");
            _pDiametre = _Config.AjouterParam("Diametre", "0", "Diametres des trous :");
            _pPercageOuvert = _Config.AjouterParam("PercageOuvert", false, "Prendre en compte les perçages ouvert");

            OnCalque += Calque;
            OnRunOkCommand += RunOkCommand;
        }

        private CtrlSelectionBox _Select_Base;
        private CtrlSelectionBox _Select_Percage;
        private CtrlTextBox _Text_Diametre;
        private CtrlCheckBox _Check_PercageOuvert;

        private CtrlCheckBox _CheckBox_EnregistrerSelection;
        private CtrlButton _Button_Preselection;

        protected void Calque()
        {
            try
            {
                Groupe G;
                G = _Calque.AjouterGroupe("Appliquer");

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

                G = _Calque.AjouterGroupe("Diametres des trous à contraindre en mm"
                                           + "\r\n  0 ou vide pour tout les perçages"
                                           + "\r\n  Valeurs séparés par une virgule");

                _Text_Diametre = G.AjouterTexteBox(_pDiametre, false);

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
            Component2 Cp = MdlBase.eSelect_RecupererComposant(1, ((CtrlSelectionBox)SelBox).Marque);
            if (Cp.IsRef() && _CheckBox_EnregistrerSelection.IsChecked)
                Param.SetValeur(Cp.eNomSansExt());
        }

        protected void PreSelection()
        {
            try
            {
                MdlBase.ClearSelection2(true);

                SelectionMgr SelMgr = MdlBase.SelectionManager;
                Component2 Piece = MdlBase.eRecChercherComposant(c => { return Regex.IsMatch(c.Name2, _pPieceBase.GetValeur<String>())
                                                                            && !c.IsSuppressed(); });
                Component2 Percage = MdlBase.eRecChercherComposant(c => { return Regex.IsMatch(c.Name2, _pPercage.GetValeur<String>())
                                                                            && !c.IsSuppressed(); });

                if (Piece.IsRef())
                    MdlBase.eSelectMulti(Piece, _Select_Percage.Marque, false);

                if (Percage.IsRef())
                    MdlBase.eSelectMulti(Percage, _Select_Percage.Marque, false);

                _Select_Base.Focus = true;

            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdInsererPercageTole Cmd = new CmdInsererPercageTole
            {
                MdlBase = MdlBase,
                CompBase = MdlBase.eSelect_RecupererComposant(1, _Select_Base.Marque),
                CompPercage = MdlBase.eSelect_RecupererComposant(1, _Select_Percage.Marque),
                ListeDiametre = new List<double>(_Text_Diametre.Text.Split(',').Select(x => { return x.Trim().eToDouble(); })),
                PercageOuvert = _Check_PercageOuvert.IsChecked
            };

            MdlBase.ClearSelection2(true);

            Cmd.Executer();
        }

    }
}
