using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ModuleCreerSymetrie
{
    [ModuleTypeDocContexte(eTypeDoc.Piece),
        ModuleTitre("Creer la symetrie de la pièce"),
        ModuleNom("CreerSymetrie"),
        ModuleDescription("Creer la symetrie de la pièce")
        ]
    public class PageCreerSymetrie : BoutonPMPManager
    {
        private Parametre PlanSymetrie;

        private ModelDoc2 MdlBase = null;
        private PartDoc Piece = null;

        public PageCreerSymetrie()
        {
            PlanSymetrie = _Config.AjouterParam("PlanSymetrie", "Plan de droite", "Selectionnez le plan de symetrie :");

            MdlBase = App.ModelDoc2;
            Piece = MdlBase.ePartDoc();

            OnCalque += Calque;
            OnRunAfterActivation += PreSelection;
            OnRunOkCommand += RunOkCommand;
        }

        private CtrlSelectionBox _Select_P_Symetrie;
        private CtrlSelectionBox _Select_Corps;

        private CtrlButton _Button_Preselection;

        protected void Calque()
        {
            try
            {
                Groupe G;
                G = _Calque.AjouterGroupe("Appliquer");

                _Button_Preselection = G.AjouterBouton("Preselectionner");
                _Button_Preselection.OnButtonPress += delegate (object sender) { PreSelection(); };

                G = _Calque.AjouterGroupe("Plan de symetrie");

                _Select_P_Symetrie = G.AjouterSelectionBox("Selectionnez le plan");
                _Select_P_Symetrie.SelectionMultipleMemeEntite = false;
                _Select_P_Symetrie.SelectionDansMultipleBox = false;
                _Select_P_Symetrie.UneSeuleEntite = true;
                _Select_P_Symetrie.FiltreSelection(swSelectType_e.swSelFACES, swSelectType_e.swSelDATUMPLANES);
                _Select_P_Symetrie.Marque = 2;

                G = _Calque.AjouterGroupe("Corps à symétriser");
                _Select_Corps = G.AjouterSelectionBox("Selectionnez les corps");
                _Select_Corps.SelectionMultipleMemeEntite = false;
                _Select_Corps.SelectionDansMultipleBox = false;
                _Select_Corps.UneSeuleEntite = false;
                _Select_Corps.Hauteur = 13;
                _Select_Corps.FiltreSelection(swSelectType_e.swSelSOLIDBODIES);
                _Select_Corps.Marque = 256;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void PreSelection()
        {
            try
            {
                MdlBase.eEffacerSelection();

                String cPlanEsquisse = PlanSymetrie.GetValeur<String>();

                Feature P = MdlBase.eChercherFonction(f => { return Regex.IsMatch(f.Name, cPlanEsquisse); }, false);
                if (P.IsRef())
                    MdlBase.eSelectMulti(P, _Select_P_Symetrie.Marque, true);

                var ListeCorps = Piece.eListeCorps();
                if (ListeCorps.Count > 0)
                    MdlBase.eSelectMulti(ListeCorps, _Select_Corps.Marque, true);
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdCreerSymetrie Cmd = new CmdCreerSymetrie();
            Cmd.MdlBase = MdlBase;
            Cmd.Plan = MdlBase.eSelect_RecupererObjet<Feature>(1, _Select_P_Symetrie.Marque);
            Cmd.ListeCorps = MdlBase.eSelect_RecupererListeObjets<Body2>(_Select_Corps.Marque);

            Cmd.Executer();
        }
    }
}
