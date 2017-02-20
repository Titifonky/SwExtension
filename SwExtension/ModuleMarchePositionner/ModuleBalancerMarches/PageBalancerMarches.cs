using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ModuleMarchePositionner
{
    namespace ModuleBalancerMarches
    {
        [ModuleTypeDocContexte(eTypeDoc.Assemblage),
            ModuleTitre("Balancer les marches"),
            ModuleNom("BalancerMarches"),
            ModuleDescription("Balancer les marches." +
                                "\r\nLa marche doit être positionnée par une contrainte d'angle"
            )
            ]
        public class PageBalancerMarches : PageMarchePositionner
        {
            public PageBalancerMarches()
            {
                try
                {
                    OnCalque += Calque;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private Groupe GroupeMarche;
            private Groupe GroupeParametres;
            private CtrlSelectionBox _Select_Marche;
            private CtrlButton _Button_Selection;

            private CtrlTextListBox _TextListBox_Marche;
            private CtrlTextBox _TextBox_Angle;

            protected void Calque()
            {
                try
                {
                    App.ModelDoc2.eEffacerSelection();

                    GroupeMarche = _Calque.AjouterGroupe("Marches");

                    _Select_Marche = GroupeMarche.AjouterSelectionBox("", "Selectionnez les marches", false);
                    _Select_Marche.SelectionMultipleMemeEntite = false;
                    _Select_Marche.SelectionDansMultipleBox = true;
                    _Select_Marche.UneSeuleEntite = false;
                    _Select_Marche.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
                    _Select_Marche.Hauteur = 7;

                    _Select_Marche.OnSubmitSelection += SelectionnerComposant1erNvx;

                    _Button_Selection = GroupeMarche.AjouterBouton("Modifier la selection");
                    _Button_Selection.OnButtonPress += MajAngle;

                    GroupeParametres = _Calque.AjouterGroupe("Parametrage");

                    _TextListBox_Marche = GroupeParametres.AjouterTextListBox("Liste des marches", "Selectionnez la marche à modifier");
                    _TextListBox_Marche.SelectionMultiple = false;
                    _TextListBox_Marche.TouteHauteur = true;
                    _TextListBox_Marche.Height = 70;
                    _TextListBox_Marche.OnSelectionChanged += SelectionChanged;

                    _TextBox_Angle = GroupeParametres.AjouterTexteBox("Angle de la marche", "");
                    _TextBox_Angle.Multiligne = false;
                    _TextBox_Angle.NotifieSurFocus = true;

                    GroupeParametres.Expanded = false;

                    GroupeMarche.OnExpand += GroupeParametres.UnExpand;
                    GroupeParametres.OnExpand += GroupeMarche.UnExpand;
                    GroupeMarche.OnExpand += Vider;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private List<Component2> ListeMarches = new List<Component2>();
            private List<String> ListeNomComp = new List<String>();
            private List<Dimension> ListeDim = new List<Dimension>();

            private void Vider()
            {
                _TextListBox_Marche.Vider();
                ListeNomComp.Clear();
                ListeDim.Clear();
                App.ModelDoc2.eSelectMulti(ListeMarches, _Select_Marche.Marque);
            }

            private void MajAngle(Object sender)
            {
                ListeMarches = App.ModelDoc2.eSelect_RecupererListeComposants(_Select_Marche.Marque);

                if (ListeMarches.Count == 0) return;

                foreach (Component2 cp in ListeMarches)
                {
                    Object[] Mates = cp.GetMates();

                    if (Mates.IsRef())
                    {
                        foreach (Mate2 mate in Mates)
                        {
                            if (mate.Type == (int)swMateType_e.swMateANGLE)
                            {
                                DisplayDimension DD = mate.DisplayDimension2[0];
                                Dimension D = DD.GetDimension2(0);
                                ListeNomComp.Add(cp.Name2);
                                ListeDim.Add(D);
                                break;
                            }
                        }
                    }
                }

                _TextListBox_Marche.Liste = ListeNomComp;
                GroupeParametres.Expanded = true;
                foreach (Component2 cp in ListeMarches)
                    cp.DeSelect();

                _TextListBox_Marche.SelectedIndex = 0;
            }

            private const int Marque = 2;
            private int LastSel = 0;

            private void SelectionChanged(Object sender, int nb)
            {
                ModifierMarche(LastSel, _TextBox_Angle.Text);
                LastSel = nb;
                ListeMarches[nb].eSelectById(App.ModelDoc2, Marque);
                _TextBox_Angle.Text = GetValDegree(ListeDim[nb]); ;
                _TextBox_Angle.Focus = true;
            }

            private void ModifierMarche(int index, String text)
            {
                Dimension dimension = ListeDim[index];

                if (dimension.IsNull()) return;

                if (GetValDegree(dimension) == text) return;

                SetValDegree(dimension, text);

                App.ModelDoc2.EditRebuild3();
            }

            private String GetValDegree(Dimension d)
            {
                Double[] tabDim = d.GetSystemValue3((int)swInConfigurationOpts_e.swThisConfiguration, null);
                Double Val = tabDim[0].Degree();
                if (Val > 180)
                    Val = (360 - Val) * -1;

                return Math.Round(Val, 2).ToString();
            }

            private void SetValDegree(Dimension d, String text)
            {
                Double Val = text.eToDouble();

                if (Val < 0)
                    Val = 360 + Val;

                d.SetSystemValue3(Val.Radian(), (int)swInConfigurationOpts_e.swThisConfiguration, null);
            }
        }
    }
}
