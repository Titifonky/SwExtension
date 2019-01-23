using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ModuleMarcheConfig
{
    namespace ModulePositionnerPlatine
    {
        [ModuleTypeDocContexte(eTypeDoc.Assemblage),
            ModuleTitre("Positionner les platines"),
            ModuleNom("PositionnerPlatine"),
            ModuleDescription("Positionne les platines suivant les bords des marches")
            ]
        public class PagePositionnerPlatine : PageMarcheConfig
        {
            private Parametre pMarche;
            private Parametre pFaceDessus;
            private Parametre pFaceDevant;
            private Parametre pPlatineG;
            private Parametre pPlatineD;
            private Parametre pPlanContrainte;
            private Parametre pToutesLesConfig;

            public PagePositionnerPlatine()
            {
                pMarche = _Config.AjouterParam("Marche", "PM", "Selectionnez la marche :");
                pFaceDessus = _Config.AjouterParam("FaceDessus", "F_Dessus", "Selectionnez la face du dessus :");
                pFaceDevant = _Config.AjouterParam("FaceDevant", "F_Devant", "Selectionnez la face de devant :");
                pPlatineG = _Config.AjouterParam("PlatineG", "PP01", "Selectionnez la platine gauche :");
                pPlatineD = _Config.AjouterParam("PlatineD", "PP02", "Selectionnez la platine droite :");
                pPlanContrainte = _Config.AjouterParam("PlanContrainte", "Plan de droite", "Nom du plan à contraindre :");

                pToutesLesConfig = _Config.AjouterParam("ToutesLesConfig", false, "Appliquer à toutes les configs");

                OnCalque += Calque;
                OnRunOkCommand += RunOkCommand;
            }

            private CtrlSelectionBox _Select_F_Dessus;
            private CtrlSelectionBox _Select_F_Devant;
            private CtrlSelectionBox _Select_PlatineG;
            private CtrlSelectionBox _Select_PlatineD;
            private CtrlCheckBox _CheckBox_ToutesLesConfig;
            private CtrlButton _Button_Preselection;

            protected void Calque()
            {
                try
                {
                    Groupe G;
                    G = _Calque.AjouterGroupe("Appliquer");

                    _CheckBox_ToutesLesConfig = G.AjouterCheckBox(pToutesLesConfig);
                    _CheckBox_EnregistrerSelection = G.AjouterCheckBox("Enregistrer les selections");
                    _Button_Preselection = G.AjouterBouton("Preselectionner");
                    _Button_Preselection.OnButtonPress += delegate (object sender) { PreSelection(); };

                    G = _Calque.AjouterGroupe("Face du dessus" + " ( " + pFaceDessus.GetValeur<String>() + "@" + pMarche.GetValeur<String>() + " )");

                    _Select_F_Dessus = G.AjouterSelectionBox("", "Selectionnez la face du dessus");
                    _Select_F_Dessus.SelectionMultipleMemeEntite = false;
                    _Select_F_Dessus.SelectionDansMultipleBox = false;
                    _Select_F_Dessus.UneSeuleEntite = true;
                    _Select_F_Dessus.FiltreSelection(swSelectType_e.swSelFACES, swSelectType_e.swSelCOMPONENTS);

                    _Select_F_Dessus.OnSubmitSelection += delegate (Object SelBox, Object selection, int selType, String itemText)
                    { return FiltreFace((CtrlSelectionBox)SelBox, selection, selType, pFaceDessus); };

                    _Select_F_Dessus.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, pMarche); };

                    G = _Calque.AjouterGroupe("Face de devant" + " ( " + pFaceDevant.GetValeur<String>() + "@" + pMarche.GetValeur<String>() + " )");

                    _Select_F_Devant = G.AjouterSelectionBox("", "Selectionnez la face de devant");
                    _Select_F_Devant.SelectionMultipleMemeEntite = false;
                    _Select_F_Devant.SelectionDansMultipleBox = false;
                    _Select_F_Devant.UneSeuleEntite = true;
                    _Select_F_Devant.FiltreSelection(swSelectType_e.swSelFACES, swSelectType_e.swSelCOMPONENTS);

                    _Select_F_Devant.OnSubmitSelection += delegate (Object SelBox, Object selection, int selType, String itemText)
                    { return FiltreFace((CtrlSelectionBox)SelBox, selection, selType, pFaceDevant); };

                    _Select_F_Devant.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, pMarche); };

                    _Select_F_Dessus.OnApplyOnSelection += _Select_F_Devant.GainedFocus;

                    G = _Calque.AjouterGroupe("Platine gauche" + " ( " + pPlanContrainte.GetValeur<String>() + "@" + pPlatineG.GetValeur<String>() + " )");

                    _Select_PlatineG = G.AjouterSelectionBox("Plan à contraindre");
                    _Select_PlatineG.SelectionMultipleMemeEntite = false;
                    _Select_PlatineG.SelectionDansMultipleBox = false;
                    _Select_PlatineG.UneSeuleEntite = true;
                    _Select_PlatineG.FiltreSelection(swSelectType_e.swSelDATUMPLANES, swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);

                    _Select_PlatineG.OnSubmitSelection += delegate (Object SelBox, Object selection, int selType, String itemText)
                    { return FiltrePlan((CtrlSelectionBox)SelBox, selection, selType, pPlanContrainte); };

                    // Svg des parametres
                    _Select_PlatineG.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, pPlatineG); };
                    _Select_PlatineG.OnApplyOnSelection += delegate (Object Box) { SvgNomFonction(Box, pPlanContrainte); };

                    _Select_F_Devant.OnApplyOnSelection += _Select_PlatineG.GainedFocus;

                    G = _Calque.AjouterGroupe("Platine droite" + " ( " + pPlanContrainte.GetValeur<String>() + "@" + pPlatineD.GetValeur<String>() + " )");

                    _Select_PlatineD = G.AjouterSelectionBox("Plan à contraindre");
                    _Select_PlatineD.SelectionMultipleMemeEntite = false;
                    _Select_PlatineD.SelectionDansMultipleBox = false;
                    _Select_PlatineD.UneSeuleEntite = true;
                    _Select_PlatineD.FiltreSelection(swSelectType_e.swSelDATUMPLANES, swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);

                    _Select_PlatineD.OnSubmitSelection += delegate (Object SelBox, Object selection, int selType, String itemText)
                    { return FiltrePlan((CtrlSelectionBox)SelBox, selection, selType, pPlanContrainte); };

                    // Svg des parametres
                    _Select_PlatineD.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, pPlatineD); };
                    _Select_PlatineD.OnApplyOnSelection += delegate (Object Box) { SvgNomFonction(Box, pPlanContrainte); };

                    _Select_PlatineG.OnApplyOnSelection += _Select_PlatineD.GainedFocus;

                    // OnCheck, on enregistre les parametres
                    _CheckBox_EnregistrerSelection.OnCheck += _Select_F_Dessus.ApplyOnSelection;
                    _CheckBox_EnregistrerSelection.OnCheck += _Select_F_Devant.ApplyOnSelection;
                    _CheckBox_EnregistrerSelection.OnCheck += _Select_PlatineG.ApplyOnSelection;
                    _CheckBox_EnregistrerSelection.OnCheck += _Select_PlatineD.ApplyOnSelection;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            protected void PreSelection()
            {
                try
                {
                    MdlBase.ClearSelection2(true);

                    Component2 Marche = MdlBase.eRecChercherComposant(c =>
                    {
                        return Regex.IsMatch(c.Name2, pMarche.GetValeur<String>())
                                 && !c.IsSuppressed();
                    });

                    if (Marche.IsRef())
                    {
                        SelectFace(_Select_F_Dessus, Marche, pFaceDessus);
                        SelectFace(_Select_F_Devant, Marche, pFaceDevant);
                    }

                    Component2 PlatineG = MdlBase.eRecChercherComposant(c =>
                    {
                        return Regex.IsMatch(c.Name2, pPlatineG.GetValeur<String>())
                       && !c.IsSuppressed();
                    });

                    if (PlatineG.IsRef())
                        SelectPlan(_Select_PlatineG, PlatineG, pPlanContrainte);


                    Component2 PlatineD = MdlBase.eRecChercherComposant(c =>
                    {
                        return Regex.IsMatch(c.Name2, pPlatineD.GetValeur<String>())
                       && !c.IsSuppressed();
                    });

                    if (PlatineD.IsRef())
                        SelectPlan(_Select_PlatineD, PlatineD, pPlanContrainte);
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            protected void RunOkCommand()
            {
                CmdPositionnerPlatine Cmd = new CmdPositionnerPlatine();
                Cmd.MdlBase = MdlBase;
                Cmd.F_Dessus = MdlBase.eSelect_RecupererObjet<Face2>(1, _Select_F_Dessus.Marque);
                Cmd.F_Devant = MdlBase.eSelect_RecupererObjet<Face2>(1, _Select_F_Devant.Marque);
                Cmd.PltG = MdlBase.eSelect_RecupererComposant(1, _Select_PlatineG.Marque);
                Cmd.PltD = MdlBase.eSelect_RecupererComposant(1, _Select_PlatineD.Marque);
                Cmd.PltG_Plan = MdlBase.eSelect_RecupererObjet<Feature>(1, _Select_PlatineG.Marque);
                Cmd.PltD_Plan = MdlBase.eSelect_RecupererObjet<Feature>(1, _Select_PlatineD.Marque);
                Cmd.SurTouteLesConfigs = _CheckBox_ToutesLesConfig.IsChecked;

                MdlBase.ClearSelection2(true);

                Cmd.Executer();
            }

        }
    }
}
