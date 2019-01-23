using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ModuleMarcheConfig
{
    namespace ModuleConfigurerContreMarche
    {
        [ModuleTypeDocContexte(eTypeDoc.Assemblage),
            ModuleTitre("Configurer les contremarches"),
            ModuleNom("ConfigurerContreMarche"),
            ModuleDescription("Configurer les contremarches" +
                              "\r\n - Nom des dimensions parametrées : Lg1, Ag1, Lc1, Lg2, Ag2, Lc2")
            ]
        public class PageConfigurerContreMarche : PageMarcheConfig
        {
            private Parametre pMarche;
            private Parametre pFaceDessus;
            private Parametre pFaceDevant;
            private Parametre pEsquisse;
            private Parametre pNomEsquisse;
            private Parametre pToutesLesConfig;

            public PageConfigurerContreMarche()
            {
                pMarche = _Config.AjouterParam("Marche", "PM", "Selectionnez la marche :");
                pFaceDessus = _Config.AjouterParam("FaceDessus", "F_Dessus", "Selectionnez la face du dessus :");
                pFaceDevant = _Config.AjouterParam("FaceDevant", "F_Devant", "Selectionnez la face de devant :");
                pEsquisse = _Config.AjouterParam("Esquisse", "PC01", "Selectionnez la contremarche :");
                pNomEsquisse = _Config.AjouterParam("NomEsquisse", "Config", "Esquisse à configurer :");

                pToutesLesConfig = _Config.AjouterParam("ToutesLesConfig", false, "Appliquer à toutes les configs");

                OnCalque += Calque;
                OnRunOkCommand += RunOkCommand;
            }

            private CtrlSelectionBox _Select_F_Dessus;
            private CtrlSelectionBox _Select_F_Devant;
            private CtrlSelectionBox _Select_ContreMarche_Esquisse;

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

                    _Select_F_Dessus = G.AjouterSelectionBox("Selectionnez la face du dessus");
                    _Select_F_Dessus.SelectionMultipleMemeEntite = false;
                    _Select_F_Dessus.SelectionDansMultipleBox = false;
                    _Select_F_Dessus.UneSeuleEntite = true;
                    _Select_F_Dessus.FiltreSelection(swSelectType_e.swSelFACES);

                    _Select_F_Dessus.OnSubmitSelection += delegate (Object SelBox, Object selection, int selType, String itemText)
                    { return FiltreFace((CtrlSelectionBox)SelBox, selection, selType, pFaceDessus); };

                    _Select_F_Dessus.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, pMarche); };

                    G = _Calque.AjouterGroupe("Face de devant" + " ( " + pFaceDevant.GetValeur<String>() + "@" + pMarche.GetValeur<String>() + " )");

                    _Select_F_Devant = G.AjouterSelectionBox("Selectionnez la face de devant");
                    _Select_F_Devant.SelectionMultipleMemeEntite = false;
                    _Select_F_Devant.SelectionDansMultipleBox = false;
                    _Select_F_Devant.UneSeuleEntite = true;
                    _Select_F_Devant.FiltreSelection(swSelectType_e.swSelFACES);

                    _Select_F_Devant.OnSubmitSelection += delegate (Object SelBox, Object selection, int selType, String itemText)
                    { return FiltreFace((CtrlSelectionBox)SelBox, selection, selType, pFaceDevant); };

                    _Select_F_Devant.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, pMarche); };

                    _Select_F_Dessus.OnApplyOnSelection += _Select_F_Devant.GainedFocus;

                    G = _Calque.AjouterGroupe("Contremarche" + " ( " + pNomEsquisse.GetValeur<String>() + "@" + pEsquisse.GetValeur<String>() + " )");

                    _Select_ContreMarche_Esquisse = G.AjouterSelectionBox("Selectionnez l'esquisse à configurer", "Esquisse à configurer");
                    _Select_ContreMarche_Esquisse.SelectionMultipleMemeEntite = false;
                    _Select_ContreMarche_Esquisse.SelectionDansMultipleBox = false;
                    _Select_ContreMarche_Esquisse.UneSeuleEntite = true;
                    _Select_ContreMarche_Esquisse.FiltreSelection(swSelectType_e.swSelSKETCHES, swSelectType_e.swSelCOMPONENTS);

                    _Select_ContreMarche_Esquisse.OnSubmitSelection += delegate (Object SelBox, Object selection, int selType, String itemText)
                    { return FiltreEsquisse((CtrlSelectionBox)SelBox, selection, selType, pNomEsquisse); };

                    _Select_ContreMarche_Esquisse.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, pEsquisse); };
                    _Select_ContreMarche_Esquisse.OnApplyOnSelection += delegate (Object Box) { SvgNomFonction(Box, pNomEsquisse); };

                    _Select_F_Devant.OnApplyOnSelection += _Select_ContreMarche_Esquisse.GainedFocus;

                    //G = _Calque.AjouterGroupe("Options");

                    //_Text_NomEsquisse = G.AjouterTexteBox(_Config.GetParam("NomEsquisse"));
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

                    Component2 ContreMarche = MdlBase.eRecChercherComposant(c =>
                    {
                        return Regex.IsMatch(c.Name2, pEsquisse.GetValeur<String>())
                       && !c.IsSuppressed();
                    });

                    if (ContreMarche.IsRef())
                        SelectEsquisse(_Select_ContreMarche_Esquisse, ContreMarche, pNomEsquisse);
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            protected void RunOkCommand()
            {
                CmdConfigurerContreMarche Cmd = new CmdConfigurerContreMarche();
                Cmd.MdlBase = MdlBase;
                Cmd.F_Dessus = MdlBase.eSelect_RecupererObjet<Face2>(1, _Select_F_Dessus.Marque);
                Cmd.F_Devant = MdlBase.eSelect_RecupererObjet<Face2>(1, _Select_F_Devant.Marque);

                Cmd.ContreMarche_Esquisse_Comp = MdlBase.eSelect_RecupererComposant(1, _Select_ContreMarche_Esquisse.Marque);
                Cmd.ContreMarche_Esquisse_Fonction = MdlBase.eSelect_RecupererObjet<Feature>(1, _Select_ContreMarche_Esquisse.Marque);

                //Cmd.NomEsquisse = _Text_NomEsquisse.Text;
                Cmd.SurTouteLesConfigs = _CheckBox_ToutesLesConfig.IsChecked;

                MdlBase.ClearSelection2(true);

                Cmd.Executer();
            }

        }
    }
}
