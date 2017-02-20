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
    namespace ModuleConfigurerPlatine
    {
        [ModuleTypeDocContexte(eTypeDoc.Assemblage),
            ModuleTitre("Configurer les platines"),
            ModuleNom("ConfigurerPlatine"),
            ModuleDescription("Configurer les platines préalablement positionnées" +
                              "\r\n - Nom des dimensions parametrées : Ag1, Ag2, Lg" +
                              "\r\n - L'origine de la platine doit être contrainte sur une face" +
                              "\r\n   attenante à la face du dessus")
            ]
        public class PageConfigurerPlatine : PageMarcheConfig
        {
            private Parametre pMarche;
            private Parametre pFaceDessus;
            private Parametre pPlatineG;
            private Parametre pPlatineD;
            private Parametre pEsquisseG;
            private Parametre pEsquisseD;
            private Parametre pPlanContrainte;
            private Parametre pNomEsquisse;
            private Parametre pLgMin;
            private Parametre pToutesLesConfig;

            public PageConfigurerPlatine()
            {
                try
                {
                    pMarche = _Config.AjouterParam("Marche", "PM", "Selectionnez la marche");
                    pFaceDessus = _Config.AjouterParam("FaceDessus", "F_Dessus", "Selectionnez la face du dessus");
                    pPlatineG = _Config.AjouterParam("PlatineG", "PP01", "Selectionnez la platine gauche");
                    pPlatineD = _Config.AjouterParam("PlatineD", "PP02", "Selectionnez la platine droite");
                    pEsquisseG = _Config.AjouterParam("EsquisseG", "PP01", "Selectionnez l'esquisse à configurer gauche");
                    pEsquisseD = _Config.AjouterParam("EsquisseD", "PP02", "Selectionnez l'esquisse à configurer droite");
                    pPlanContrainte = _Config.AjouterParam("PlanContrainte", "Plan de droite", "Nom du plan à contraindre");
                    pNomEsquisse = _Config.AjouterParam("NomEsquisse", "Config", "Esquisse à configurer");
                    pLgMin = _Config.AjouterParam("LgMin", 50.0, "Lg mini de la platine (mm)");

                    pToutesLesConfig = _Config.AjouterParam("ToutesLesConfig", false, "Appliquer à toutes les configs");

                    OnCalque += Calque;
                    OnRunOkCommand += RunOkCommand;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private CtrlSelectionBox _Select_F_Dessus;
            private CtrlSelectionBox _Select_PlatineG;
            private CtrlSelectionBox _Select_PlatineG_Esquisse;
            private CtrlSelectionBox _Select_PlatineD;
            private CtrlSelectionBox _Select_PlatineD_Esquisse;

            private CtrlTextBox _Text_LgMini;
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
                    _Select_F_Dessus.FiltreSelection(swSelectType_e.swSelFACES);

                    _Select_F_Dessus.OnSubmitSelection += delegate (Object SelBox, Object selection, int selType, String itemText)
                    { return FiltreFace((CtrlSelectionBox)SelBox, selection, selType, pFaceDessus); };

                    _Select_F_Dessus.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, pMarche); };

                    G = _Calque.AjouterGroupe("Platine gauche");

                    _Select_PlatineG = G.AjouterSelectionBox("Plan" + " ( " + pPlanContrainte.GetValeur<String>() + "@" + pPlatineG.GetValeur<String>() + " )", "Selectionnez le plan");
                    _Select_PlatineG.SelectionMultipleMemeEntite = false;
                    _Select_PlatineG.SelectionDansMultipleBox = false;
                    _Select_PlatineG.UneSeuleEntite = true;
                    _Select_PlatineG.FiltreSelection(swSelectType_e.swSelDATUMPLANES, swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);

                    _Select_PlatineG.OnSubmitSelection += delegate (Object SelBox, Object selection, int selType, String itemText)
                    { return FiltrePlan((CtrlSelectionBox)SelBox, selection, selType, pPlanContrainte); };

                    // Svg des parametres
                    _Select_PlatineG.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, pPlatineG); };
                    _Select_PlatineG.OnApplyOnSelection += delegate (Object Box) { SvgNomFonction(Box, pPlanContrainte); };

                    _Select_F_Dessus.OnApplyOnSelection += _Select_PlatineG.GainedFocus;

                    _Select_PlatineG_Esquisse = G.AjouterSelectionBox("Esquisse" + " ( " + pNomEsquisse.GetValeur<String>() + "@" + pEsquisseG.GetValeur<String>() + " )", "Selectionnez l'esquisse");
                    _Select_PlatineG_Esquisse.SelectionMultipleMemeEntite = false;
                    _Select_PlatineG_Esquisse.SelectionDansMultipleBox = false;
                    _Select_PlatineG_Esquisse.UneSeuleEntite = true;
                    _Select_PlatineG_Esquisse.FiltreSelection(swSelectType_e.swSelSKETCHES, swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);

                    _Select_PlatineG_Esquisse.OnSubmitSelection += delegate (Object SelBox, Object selection, int selType, String itemText)
                    { return FiltreEsquisse((CtrlSelectionBox)SelBox, selection, selType, pNomEsquisse); };

                    _Select_PlatineG_Esquisse.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, pEsquisseG); };
                    _Select_PlatineG_Esquisse.OnApplyOnSelection += delegate (Object Box) { SvgNomFonction(Box, pNomEsquisse); };

                    _Select_PlatineG.OnApplyOnSelection += _Select_PlatineG_Esquisse.GainedFocus;

                    G = _Calque.AjouterGroupe("Platine droite");

                    _Select_PlatineD = G.AjouterSelectionBox("Plan" + " ( " + pPlanContrainte.GetValeur<String>() + "@" + pPlatineD.GetValeur<String>() + " )", "Selectionnez le plan");
                    _Select_PlatineD.SelectionMultipleMemeEntite = false;
                    _Select_PlatineD.SelectionDansMultipleBox = false;
                    _Select_PlatineD.UneSeuleEntite = true;
                    _Select_PlatineD.FiltreSelection(swSelectType_e.swSelDATUMPLANES, swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);

                    _Select_PlatineD.OnSubmitSelection += delegate (Object SelBox, Object selection, int selType, String itemText)
                    { return FiltrePlan((CtrlSelectionBox)SelBox, selection, selType, pPlanContrainte); };

                    _Select_PlatineD.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, pPlatineD); };
                    _Select_PlatineD.OnApplyOnSelection += delegate (Object Box) { SvgNomFonction(Box, pPlanContrainte); };

                    _Select_PlatineG_Esquisse.OnApplyOnSelection += _Select_PlatineD.GainedFocus;

                    _Select_PlatineD_Esquisse = G.AjouterSelectionBox("Esquisse" + " ( " + pNomEsquisse.GetValeur<String>() + "@" + pEsquisseD.GetValeur<String>() + " )", "Selectionnez l'esquisse");
                    _Select_PlatineD_Esquisse.SelectionMultipleMemeEntite = false;
                    _Select_PlatineD_Esquisse.SelectionDansMultipleBox = false;
                    _Select_PlatineD_Esquisse.UneSeuleEntite = true;
                    _Select_PlatineD_Esquisse.FiltreSelection(swSelectType_e.swSelSKETCHES, swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);

                    _Select_PlatineD_Esquisse.OnSubmitSelection += delegate (Object SelBox, Object selection, int selType, String itemText)
                    { return FiltreEsquisse((CtrlSelectionBox)SelBox, selection, selType, pNomEsquisse); };

                    _Select_PlatineD_Esquisse.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, pEsquisseD); };
                    _Select_PlatineD_Esquisse.OnApplyOnSelection += delegate (Object Box) { SvgNomFonction(Box, pNomEsquisse); };

                    _Select_PlatineD.OnApplyOnSelection += _Select_PlatineD_Esquisse.GainedFocus;

                    G = _Calque.AjouterGroupe("Options");

                    _Text_LgMini = G.AjouterTexteBox(pLgMin, true);

                    // OnCheck, on enregistre les parametres
                    _CheckBox_EnregistrerSelection.OnCheck += _Select_F_Dessus.ApplyOnSelection;
                    _CheckBox_EnregistrerSelection.OnCheck += _Select_PlatineG.ApplyOnSelection;
                    _CheckBox_EnregistrerSelection.OnCheck += _Select_PlatineG_Esquisse.ApplyOnSelection;
                    _CheckBox_EnregistrerSelection.OnCheck += _Select_PlatineD.ApplyOnSelection;
                    _CheckBox_EnregistrerSelection.OnCheck += _Select_PlatineD_Esquisse.ApplyOnSelection;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            protected void PreSelection()
            {
                try
                {
                    App.ModelDoc2.ClearSelection2(true);

                    Component2 Marche = App.ModelDoc2.eRecChercherComposant(c =>
                    {
                        return Regex.IsMatch(c.Name2, pMarche.GetValeur<String>())
                     && !c.IsSuppressed();
                    });

                    if (Marche.IsRef())
                        SelectFace(_Select_F_Dessus, Marche, pFaceDessus);

                    Component2 PlatineG = App.ModelDoc2.eRecChercherComposant(c =>
                    {
                        return Regex.IsMatch(c.Name2, pPlatineG.GetValeur<String>())
                       && !c.IsSuppressed();
                    });

                    if (PlatineG.IsRef())
                    {
                        SelectPlan(_Select_PlatineG, PlatineG, pPlanContrainte);

                        Component2 EsquisseG = App.ModelDoc2.eRecChercherComposant(c =>
                        {
                            return Regex.IsMatch(c.Name2, pEsquisseG.GetValeur<String>())
                          && !c.IsSuppressed();
                        });

                        if (!SelectEsquisse(_Select_PlatineG_Esquisse, EsquisseG, pNomEsquisse))
                        {
                            EsquisseG = PlatineG.eRecChercherComposant(c =>
                            {
                                return Regex.IsMatch(c.Name2, pEsquisseG.GetValeur<String>())
                                      && !c.IsSuppressed();
                            });

                            SelectEsquisse(_Select_PlatineG_Esquisse, EsquisseG, pNomEsquisse);
                        }
                    }

                    Component2 PlatineD = App.ModelDoc2.eRecChercherComposant(c =>
                    {
                        return Regex.IsMatch(c.Name2, pPlatineD.GetValeur<String>())
                       && !c.IsSuppressed();
                    });


                    if (PlatineD.IsRef())
                    {
                        SelectPlan(_Select_PlatineD, PlatineD, pPlanContrainte);

                        Component2 EsquisseD = App.ModelDoc2.eRecChercherComposant(c =>
                        {
                            return Regex.IsMatch(c.Name2, pEsquisseD.GetValeur<String>())
                          && !c.IsSuppressed();
                        });

                        if (!SelectEsquisse(_Select_PlatineD_Esquisse, EsquisseD, pNomEsquisse))
                        {
                            EsquisseD = PlatineD.eRecChercherComposant(c =>
                            {
                                return Regex.IsMatch(c.Name2, pEsquisseD.GetValeur<String>())
                                      && !c.IsSuppressed();
                            });

                            SelectEsquisse(_Select_PlatineD_Esquisse, EsquisseD, pNomEsquisse);
                        }
                    }
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            protected void RunOkCommand()
            {
                CmdConfigurerPlatine Cmd = new CmdConfigurerPlatine();
                Cmd.MdlBase = App.ModelDoc2;
                Cmd.F_Dessus = App.ModelDoc2.eSelect_RecupererObjet<Face2>(1, _Select_F_Dessus.Marque);
                Cmd.Marche = App.ModelDoc2.eSelect_RecupererComposant(1, _Select_F_Dessus.Marque);
                Cmd.PltG_Contrainte_Comp = App.ModelDoc2.eSelect_RecupererComposant(1, _Select_PlatineG.Marque);
                Cmd.PltD_Contrainte_Comp = App.ModelDoc2.eSelect_RecupererComposant(1, _Select_PlatineD.Marque);
                Cmd.PltG_Contrainte_Plan = App.ModelDoc2.eSelect_RecupererObjet<Feature>(1, _Select_PlatineG.Marque);
                Cmd.PltD_Contrainte_Plan = App.ModelDoc2.eSelect_RecupererObjet<Feature>(1, _Select_PlatineD.Marque);

                Cmd.PltG_Esquisse_Comp = App.ModelDoc2.eSelect_RecupererComposant(1, _Select_PlatineG_Esquisse.Marque);
                Cmd.PltD_Esquisse_Comp = App.ModelDoc2.eSelect_RecupererComposant(1, _Select_PlatineD_Esquisse.Marque);
                Cmd.PltG_Esquisse_Fonction = App.ModelDoc2.eSelect_RecupererObjet<Feature>(1, _Select_PlatineG_Esquisse.Marque);
                Cmd.PltD_Esquisse_Fonction = App.ModelDoc2.eSelect_RecupererObjet<Feature>(1, _Select_PlatineD_Esquisse.Marque);

                Cmd.LgMin = _Text_LgMini.GetTextAs<Double>() * 0.001; // On converti en metres
                Cmd.SurTouteLesConfigs = _CheckBox_ToutesLesConfig.IsChecked;

                App.ModelDoc2.ClearSelection2(true);

                Cmd.Executer();
            }
        }
    }
}
