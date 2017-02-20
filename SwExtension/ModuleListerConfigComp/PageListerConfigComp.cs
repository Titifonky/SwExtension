using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace ModuleListerConfigComp
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage),
        ModuleTitre("Lister les configs"),
        ModuleNom("ListerConfigComp"),
        ModuleDescription("Lister les configurations d'un composant.")
        ]
    public class PageListerConfigComp : BoutonPMPManager
    {
        private const String NomVolume = "Volume";
        private const int Marque = 4;

        public PageListerConfigComp()
        {
            LogToWindowLog = false;

            try
            {
                OnCalque += Calque;
                OnRunAfterActivation += AfficherConfigs;
                OnRunAfterClose += ExitIsoler;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private CtrlSelectionBox _Select_CompBase;
        private CtrlTextBox _TextBox_NomComp;
        
        private CtrlTextListBox _TextListBox_Configs;
        private CtrlSelectionBox _Select_Configs;
        private CtrlCheckBox _CheckBox_IsolerComposants;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Options");

                _CheckBox_IsolerComposants = G.AjouterCheckBox("Isoler les composants selectionnés");
                _CheckBox_IsolerComposants.OnIsCheck += delegate (Object sender, Boolean value) { SelectionChanged(null, _TextListBox_Configs.SelectedIndex); };

                G = _Calque.AjouterGroupe("Composant");

                _Select_CompBase = G.AjouterSelectionBox("", "Selectionnez le composant");
                _Select_CompBase.SelectionMultipleMemeEntite = true;
                _Select_CompBase.SelectionDansMultipleBox = true;
                _Select_CompBase.UneSeuleEntite = true;
                _Select_CompBase.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
                _Select_CompBase.OnSelectionChanged += delegate (Object sender, int nb) { AfficherConfigs(); };
                _Select_CompBase.Hauteur = 2;
                _Select_CompBase.Focus = true;

                _TextBox_NomComp = G.AjouterTexteBox("");
                _TextBox_NomComp.LectureSeule = true;

                _TextListBox_Configs = G.AjouterTextListBox("Liste des configurations dans le modèle");
                _TextListBox_Configs.TouteHauteur = true;
                _TextListBox_Configs.Height = 80;
                _TextListBox_Configs.SelectionMultiple = false;
                _TextListBox_Configs.OnSelectionChanged += SelectionChanged;

                G = _Calque.AjouterGroupe("Configurations du composant");

                _Select_Configs = G.AjouterSelectionBox("");
                _Select_Configs.SelectionMultipleMemeEntite = true;
                _Select_Configs.SelectionDansMultipleBox = true;
                _Select_Configs.UneSeuleEntite = false;
                _Select_Configs.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
                _Select_Configs.Hauteur = 15;

            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private Component2 CompBase = null;

        private void ExitIsoler()
        {
            App.Assembly.ExitIsolate();
        }

        private void IsolerComposants()
        {
            var mdl = App.ModelDoc2;

            List<Component2> ListeCompBase = App.ModelDoc2.eSelect_RecupererListeObjets<Component2>(_Select_Configs.Marque);

            mdl.eAssemblyDoc().Isolate();

            mdl.eSelectMulti(ListeCompBase, _Select_Configs.Marque, false);
        }

        private void SelectionChanged(Object sender, int Item)
        {
            var mdl = App.ModelDoc2;

            var TextListBox = sender as CtrlTextListBox;

            List<Component2> listeComps = Bdd.ListeComposants(TextListBox.SelectedText);

            ExitIsoler();

            mdl.eSelectMulti(listeComps, _Select_Configs.Marque, false);

            if (_CheckBox_IsolerComposants.IsChecked)
            {
                IsolerComposants();
                mdl.eSelectMulti(listeComps, _Select_Configs.Marque, false);
            }
        }

        private void AfficherConfigs()
        {
            var Comp = App.ModelDoc2.eSelect_RecupererComposant(1, _Select_CompBase.Marque);

            if (Comp.IsRef())
            {
                if (CompBase.IsRef() && (CompBase.GetPathName() == Comp.GetPathName()))
                    return;

                Comp.eDeSelectById(App.ModelDoc2);
                CompBase = Comp;
                _TextListBox_Configs.Vider();
            }
            else
                return;

            _TextBox_NomComp.Text = Comp.eNomSansExt();

            Rechercher_Composants(CompBase);
            _TextListBox_Configs.Liste = Bdd.ListeNomsConfigs();
        }

        private BDD Bdd;

        private void Rechercher_Composants(Component2 compBase)
        {
            Bdd = new BDD();

            var mdl = App.ModelDoc2;

            App.ModelDoc2.eRecParcourirComposants(
                    c =>
                    {
                        if (!c.IsSuppressed() && (c.GetPathName() == compBase.GetPathName()))
                            Bdd.AjouterComposant(c);

                        return false;
                    }
                );
        }

        public class BDD
        {
            private Dictionary<String, List<Component2>> Dic = new Dictionary<String, List<Component2>>();

            public void AjouterComposant(Component2 comp)
            {
                if (Dic.ContainsKey(comp.ReferencedConfiguration))
                    Dic[comp.ReferencedConfiguration].Add(comp);
                else
                    Dic.Add(comp.ReferencedConfiguration, new List<Component2>() { comp });
            }

            private Dictionary<String, String> _DicIntitule = new Dictionary<String, String>();

            public List<String> ListeNomsConfigs()
            {
                int lgInt = 0;
                int lgNb = 0;
                foreach (var Cfg in Dic.Keys)
                {
                    lgInt = Math.Max(lgInt, Cfg.Length);
                    lgNb = Math.Max(lgNb, Dic[Cfg].Count.ToString().Length);
                }

                String format = "{0,-" + (lgInt + 2) + "}  × {1,-" + lgNb + "}";

                foreach (var Cfg in Dic.Keys)
                {
                    _DicIntitule.Add(String.Format(format, Cfg, Dic[Cfg].Count), Cfg);
                }

                return _DicIntitule.Keys.ToList();
            }

            public List<Component2> ListeComposants(String Intitule)
            {
                return Dic[_DicIntitule[Intitule]];
            }
        }
    }
}
