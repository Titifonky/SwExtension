using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ModuleEmpreinte
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage),
        ModuleTitre("Empreinte entre les composants"),
        ModuleNom("Empreinte"),
        ModuleDescription("Empreinte entre les composants sélectionnés")
        ]
    public class PageEmpreinte : BoutonPMPManager
    {
        private Parametre PrefixeBase;
        private Parametre PrefixeEmpreinte;
        private Parametre PropEmpreinte;
        private Parametre PropPrefixeEmpreinte;
        private Parametre PropTarauderEmpreinte;

        public PageEmpreinte()
        {
            PrefixeBase = _Config.AjouterParam("PrefixeBase", "", "Filtrer les prefixes (?*#):");
            PrefixeEmpreinte = _Config.AjouterParam("PrefixeEmpreinte", "", "Filtrer les prefixes (?*#):");
            PropEmpreinte = _Config.AjouterParam("PropEmpreinte", "Empreinte", "Propriete \"Empreinte\"");
            PropPrefixeEmpreinte = _Config.AjouterParam("PropPrefixeEmpreinte", "PrefixeEmpreinte", "Propriete \"PrefixeEmpreinte\"");
            PropTarauderEmpreinte = _Config.AjouterParam("PropTarauderEmpreinte", "TarauderEmpreinte", "Propriete \"TarauderEmpreinte\"");

            Empreinte.NomPropEmpreinte = PropEmpreinte.GetValeur<String>();
            Empreinte.NomPropPrefixe = PropPrefixeEmpreinte.GetValeur<String>();

            OnCalque += Calque;
            OnRunOkCommand += RunOkCommand;
            OnRunAfterActivation += delegate { _FiltreCompBase.AppliquerParam(); _FiltreCompEmpreinte.AppliquerParam(); };
            OnRunAfterClose += RunAfterClose;
        }

        private Groupe _Groupe1;
        private Groupe _Groupe2;

        private CtrlSelectionBox _Select_CompBase;
        private CtrlSelectionBox _Select_CompEmpreinte;

        private FiltreComp _FiltreCompBase;
        private FiltreComp _FiltreCompEmpreinte;
        private CtrlCheckBox _CheckBox_MasquerLesEmpreintes;

        private CtrlButton _Button_IsolerComposants;

        protected void Calque()
        {
            try
            {
                Groupe G;

                _Groupe1 = _Calque.AjouterGroupe("Selectionner les composants de base");

                _Select_CompBase = _Groupe1.AjouterSelectionBox("", "Selectionnez les composants");
                _Select_CompBase.SelectionMultipleMemeEntite = false;
                _Select_CompBase.SelectionDansMultipleBox = false;
                _Select_CompBase.UneSeuleEntite = false;
                _Select_CompBase.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
                _Select_CompBase.OnSubmitSelection += SelectionnerPiece;
                _Select_CompBase.Hauteur = 8;
                _Select_CompBase.Focus = true;

                Isoler.ListSelectionBox.Add(_Select_CompBase);

                _FiltreCompBase = new FiltreComp(MdlBase, _Groupe1, _Select_CompBase, PrefixeBase);

                _Groupe2 = _Calque.AjouterGroupe("Selectionner les composants empreinte");

                _Select_CompEmpreinte = _Groupe2.AjouterSelectionBox("", "Selectionnez les composants");
                _Select_CompEmpreinte.SelectionMultipleMemeEntite = false;
                _Select_CompEmpreinte.SelectionDansMultipleBox = false;
                _Select_CompEmpreinte.UneSeuleEntite = false;
                _Select_CompEmpreinte.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
                _Select_CompEmpreinte.OnSubmitSelection += SelectionnerPiece;
                _Select_CompEmpreinte.Hauteur = 8;

                Isoler.ListSelectionBox.Add(_Select_CompEmpreinte);

                _FiltreCompEmpreinte = new FiltreComp(MdlBase, _Groupe2, _Select_CompEmpreinte, PrefixeEmpreinte);

                _Groupe1.OnExpand += _Groupe2.UnExpand;
                _Groupe2.OnExpand += _Groupe1.UnExpand;
                _Groupe1.Expand();

                G = _Calque.AjouterGroupe("Options");

                _Button_IsolerComposants = G.AjouterBouton("Isoler les composants");
                _Button_IsolerComposants.OnButtonPress += delegate (Object sender) { Isoler.Run(MdlBase); };

                Isoler.Bouton = _Button_IsolerComposants;

                _CheckBox_MasquerLesEmpreintes = G.AjouterCheckBox("Masquer toutes les empreintes");

            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private class FiltreComp
        {
            public Groupe G;

            private CtrlSelectionBox _SelectionBox;

            private CtrlSelectionBox _Select_CompBase;
            private CtrlTextBox _TextBox_NomComp;

            private CtrlOption o1;
            private CtrlOption o2;

            private CtrlTextBox _TextBox_Prefixe;

            private CtrlTextListBox _TextListBox_Configs;

            private CtrlButton _Button;

            private ModelDoc2 _Mdl;

            public FiltreComp(ModelDoc2 mdl, Groupe g, CtrlSelectionBox SelectionBox, Parametre Prefixe)
            {
                _Mdl = mdl;
                _SelectionBox = SelectionBox;

                G = g;

                _Select_CompBase = G.AjouterSelectionBox("", "Selectionnez le composant");
                _Select_CompBase.SelectionMultipleMemeEntite = true;
                _Select_CompBase.SelectionDansMultipleBox = true;
                _Select_CompBase.UneSeuleEntite = true;
                _Select_CompBase.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
                _Select_CompBase.OnSelectionChanged += delegate (Object sender, int nb) { AfficherInfos(); };
                _Select_CompBase.OnSubmitSelection += SelectionnerPiece;
                _Select_CompBase.Hauteur = 2;

                _TextBox_NomComp = G.AjouterTexteBox("");
                _TextBox_NomComp.LectureSeule = true;
                _TextBox_NomComp.NotifieSurFocus = true;

                o1 = G.AjouterOption("Filtrer propriete");
                o2 = G.AjouterOption("Filtrer config");

                o1.OnCheck += delegate (Object sender) { _TextBox_Prefixe.Visible = true; _TextListBox_Configs.Visible = false; };
                o2.OnCheck += delegate (Object sender) { _TextBox_Prefixe.Visible = false; _TextListBox_Configs.Visible = true; };

                _TextBox_Prefixe = G.AjouterTexteBox(Prefixe, true);

                _TextListBox_Configs = G.AjouterTextListBox("Liste des configurations");
                _TextListBox_Configs.TouteHauteur = true;
                _TextListBox_Configs.Height = 50;
                _TextListBox_Configs.SelectionMultiple = true;

                _Button = G.AjouterBouton("Filtrer");
                _Button.OnButtonPress += delegate (Object sender)
                {
                    if (o1.IsChecked)
                        RechercherComp(_SelectionBox, _TextBox_Prefixe.Text);
                    else
                        SelectionnerComposants();
                };
            }

            public void AppliquerParam()
            {
                o1.IsChecked = true;
            }

            private void AfficherInfos()
            {
                var Comp = _Mdl.eSelect_RecupererComposant(1, _Select_CompBase.Marque);

                Comp.eDeSelectById(_Mdl);

                if (Comp.IsNull() || (CompBase.IsRef() && (CompBase.GetPathName() == Comp.GetPathName())))
                    return;

                CompBase = Comp;
                _TextListBox_Configs.Vider();
                _TextBox_NomComp.Text = Comp.eNomSansExt();

                Rechercher_Composants(CompBase);
                _TextListBox_Configs.Liste = Bdd.ListeTextBox();
                _TextListBox_Configs.SelectedIndex = 0;

                var valprefixe = Empreinte.ValProp(Comp).Trim();
                if (!String.IsNullOrWhiteSpace(valprefixe))
                    _TextBox_Prefixe.Text = valprefixe;
            }

            //==================================

            private void RechercherComp(CtrlSelectionBox box, String pattern)
            {
                try
                {
                    if (String.IsNullOrWhiteSpace(pattern))
                        return;

                    String[] listePattern = pattern.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    //{
                    //    var lcp = _Mdl.eSelect_RecupererListeObjets<Component2>(box.Marque);
                    //    foreach (Component2 c in lcp)
                    //        c.eDeSelectById(_Mdl);
                    //}

                    box.Focus = true;

                    {
                        var lcp = new List<Component2>();
                        _Mdl.eRecParcourirComposants(
                            c =>
                            {

                                if (!c.IsSuppressed() && (c.TypeDoc() == eTypeDoc.Piece))
                                {
                                    if (c.ePropExiste(Empreinte.NomPropEmpreinte) && (c.eProp(Empreinte.NomPropEmpreinte) == "1"))
                                    {
                                        if (TestStringLikeListePattern(c.eProp(Empreinte.NomPropPrefixe), listePattern))
                                            lcp.Add(c);
                                    }
                                }
                                return false;
                            },
                            null
                            );

                        Isoler.Exit(_Mdl);

                        _Mdl.eSelectMulti(lcp, box.Marque, true);
                    }
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }

            }

            private Boolean TestStringLikeListePattern(String s, String[] Liste)
            {
                foreach (var t in Liste)
                {
                    if (s.eIsLike(t))
                        return true;
                }

                return false;
            }

            //==================================

            private Component2 CompBase = null;

            private void SelectionnerComposants()
            {
                List<Component2> listeComps = Bdd.ListeComposants(_TextListBox_Configs.ListSelectedIndex);

                _Mdl.eSelectMulti(listeComps, _SelectionBox.Marque, true);
            }

            private BDD Bdd;

            private void Rechercher_Composants(Component2 compBase)
            {
                Bdd = new BDD();

                _Mdl.eRecParcourirComposants(
                        c =>
                        {
                            if (!c.IsSuppressed() && (c.GetPathName() == compBase.GetPathName()))
                                Bdd.AjouterComposant(c);

                            return false;
                        }
                    );
            }

            private class BDD
            {
                private SortedDictionary<String, List<Component2>> Dic = new SortedDictionary<String, List<Component2>>(new WindowsStringComparer());

                public void AjouterComposant(Component2 comp)
                {
                    if (Dic.ContainsKey(comp.ReferencedConfiguration))
                        Dic[comp.ReferencedConfiguration].Add(comp);
                    else
                        Dic.Add(comp.ReferencedConfiguration, new List<Component2>() { comp });
                }

                public List<String> ListeTextBox()
                {
                    List<String> Liste = new List<String>();

                    foreach (var item in Dic)
                        Liste.Add( String.Format("{0,-7} x{1,-4}", item.Key, item.Value.Count));

                    return Liste;
                }

                private List<String> _ListeNomsConfigs;
                private List<String> ListeNomsConfigs
                {
                    get
                    {
                        if (_ListeNomsConfigs.IsNull())
                            _ListeNomsConfigs = Dic.Keys.ToList();

                        return _ListeNomsConfigs;
                    }
                }

                public List<Component2> ListeComposants(List<int> listeIndex)
                {
                    List<Component2> Liste = new List<Component2>();

                    foreach (var index in listeIndex)
                        Liste.AddRange(Dic[ListeNomsConfigs[index]]);

                    return Liste;
                }

                public List<Component2> ListeComposants(int index)
                {
                    List<Component2> Liste = new List<Component2>();

                    Liste.AddRange(Dic[ListeNomsConfigs[index]]);

                    return Liste;
                }
            }
        }

        protected void RunOkCommand()
        {

            List<Component2> ListeCompBase = MdlBase.eSelect_RecupererListeComposants(_Select_CompBase.Marque);
            List<Component2> ListeCompEmpreinte = MdlBase.eSelect_RecupererListeComposants(_Select_CompEmpreinte.Marque);

            CmdEmpreinte Cmd = new CmdEmpreinte();
            Cmd.MdlBase = MdlBase;
            Cmd.ListeCompBase = ListeCompBase;
            Cmd.ListeCompEmpreinte = ListeCompEmpreinte;

            Cmd.Executer();
        }

        protected void RunAfterClose()
        {
            List<Component2> ListeCompEmpreinte = MdlBase.eSelect_RecupererListeObjets<Component2>(_Select_CompEmpreinte.Marque);

            Isoler.Exit(MdlBase);

            if (_CheckBox_MasquerLesEmpreintes.IsChecked == true)
            {
                WindowLog.Ecrire("Masque les composants");
                foreach (Component2 c in ListeCompEmpreinte)
                    c.Visible = (int)swComponentVisibilityState_e.swComponentHidden;
            }
        }
    }

    public static class Isoler
    {
        private static Boolean _Isoler = false;

        public static CtrlButton Bouton;

        public static List<CtrlSelectionBox> ListSelectionBox = new List<CtrlSelectionBox>();

        public static void Exit(ModelDoc2 mdl)
        {
            mdl.eAssemblyDoc().ExitIsolate();
            Bouton.Caption = "Isoler les composants";
            _Isoler = false;
        }

        public static void Run(ModelDoc2 mdl)
        {
            if (!_Isoler)
            {
                var dic = new Dictionary<int, List<Component2>>();

                foreach (var box in ListSelectionBox)
                {
                    int marque = box.Marque;
                    dic.Add(marque, mdl.eSelect_RecupererListeObjets<Component2>(marque));
                }

                mdl.eAssemblyDoc().Isolate();

                Boolean ajouter = false;

                foreach (var item in dic)
                {
                    mdl.eSelectMulti(item.Value, item.Key, ajouter);
                    ajouter = true;
                }

                _Isoler = true;

                Bouton.Caption = "Afficher tout les composants";
            }
            else
                Exit(mdl);
        }
    }

    public static class Empreinte
    {
        public static String NomPropEmpreinte;
        public static String NomPropPrefixe;

        public const String NOM_FONCTION = "Empreinte";

        public static String ValProp(this Component2 cp)
        {
            if (cp.ePropExiste(NomPropEmpreinte) && (cp.eProp(NomPropEmpreinte) == "1"))
            {
                return cp.eProp(NomPropPrefixe);
            }

            return "";
        }
    }
}
