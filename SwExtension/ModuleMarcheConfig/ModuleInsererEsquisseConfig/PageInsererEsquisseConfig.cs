using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Text.RegularExpressions;

namespace ModuleMarcheConfig
{
    namespace ModuleInsererEsquisseConfig
    {
        [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
            ModuleTitre("Inserer l'esquisse à configurer"),
            ModuleNom("InsererEsquisseConfig"),
            ModuleDescription("Creer l'esquisse à configurer")
            ]
        public class PageInsererEsquisseConfig : BoutonPMPManager
        {

            private Parametre _pPlanDessus;
            private Parametre _pNomEsquisse;

            public PageInsererEsquisseConfig()
            {
                _pPlanDessus = _Config.AjouterParam("PlanDessus", "Plan de dessus", "Selectionnez le plan sur lequel inserer l'esquisse :");
                _pNomEsquisse = _Config.AjouterParam("NomEsquisse", "Config", "Nom de l'esquisse à configurer :");

                OnCalque += Calque;
                OnRunOkCommand += RunOkCommand;
            }

            private CtrlSelectionBox _Select_F_Dessus;
            private CtrlTextBox _Text_NomEsquisse;

            protected void Calque()
            {
                try
                {
                    Groupe G;

                    G = _Calque.AjouterGroupe("Plan d'esquisse");

                    _Select_F_Dessus = G.AjouterSelectionBox("Selectionnez le plan");
                    _Select_F_Dessus.SelectionMultipleMemeEntite = false;
                    _Select_F_Dessus.SelectionDansMultipleBox = false;
                    _Select_F_Dessus.UneSeuleEntite = true;
                    _Select_F_Dessus.FiltreSelection(swSelectType_e.swSelFACES, swSelectType_e.swSelDATUMPLANES);

                    G = _Calque.AjouterGroupe("Nom de l'esquisse");
                    _Text_NomEsquisse = G.AjouterTexteBox(_pNomEsquisse);
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            protected void PreSelection()
            {
                try
                {
                    MdlBase.ClearSelection2(true);

                    String cPlanEsquisse = _pPlanDessus.GetValeur<String>();

                    Feature P = MdlBase.eChercherFonction(f => { return Regex.IsMatch(f.Name, cPlanEsquisse); }, false);
                    if (P.IsRef())
                        MdlBase.eSelectMulti(P, _Select_F_Dessus.Marque, true);

                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            protected void RunOkCommand()
            {
                CmdInsererEsquisseConfig Cmd = new CmdInsererEsquisseConfig();
                Cmd.MdlBase = MdlBase;
                Cmd.Plan = MdlBase.eSelect_RecupererObjet<Feature>(1, _Select_F_Dessus.Marque);
                Cmd.NomEsquisse = _Text_NomEsquisse.Text.Trim();

                MdlBase.ClearSelection2(true);

                Cmd.Executer();
            }

        }
    }
}
