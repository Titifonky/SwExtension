using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ModuleListerPercage
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage),
        ModuleTitre("Lister les percages"),
        ModuleNom("ListerPercage"),
        ModuleDescription("Lister les percages d'un modele.")
        ]
    public class PageListerPercage : BoutonPMPManager
    {
        public PageListerPercage()
        {
            LogToWindowLog = false;

            try
            {
                OnCalque += Calque;
                OnRunAfterActivation += AfficherPercage;
                OnRunOkCommand += RunOkCommand;
                OnRunAfterClose += ExitIsoler;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private CtrlTextListBox _TextListBox_Percage;
        private CtrlSelectionBox _Select_Selection;

        private CtrlButton _Button_IsolerComposants;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Options");

                _Button_IsolerComposants = G.AjouterBouton("Isoler les composants");
                _Button_IsolerComposants.OnButtonPress += delegate (Object sender) { IsolerComposants(); };

                G = _Calque.AjouterGroupe("Percages :");

                _TextListBox_Percage = G.AjouterTextListBox();
                _TextListBox_Percage.TouteHauteur = true;
                _TextListBox_Percage.Height = 170;
                _TextListBox_Percage.SelectionMultiple = false;
                _TextListBox_Percage.OnSelectionChanged += SelectionChanged;

                G = _Calque.AjouterGroupe("Selection");

                _Select_Selection = G.AjouterSelectionBox("");
                _Select_Selection.SelectionMultipleMemeEntite = true;
                _Select_Selection.SelectionDansMultipleBox = true;
                _Select_Selection.UneSeuleEntite = false;
                _Select_Selection.FiltreSelection(new List<swSelectType_e>() { swSelectType_e.swSelFACES });
                _Select_Selection.Hauteur = 15;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private HashSet<String> ListeHiddenComposants = new HashSet<string>();

        private Boolean Isoler = false;

        private void ExitIsoler()
        {
            App.Assembly.ExitIsolate();
            _Button_IsolerComposants.Caption = "Isoler les composants";
            Isoler = false;
        }

        private void IsolerComposants()
        {
            if (!Isoler)
            {
                var mdl = App.ModelDoc2;

                List<Body2> listeCorps = mdl.eSelect_RecupererListeObjets<Body2>(_Select_Selection.Marque);
                List<Component2> listeComps = mdl.eSelect_RecupererListeComposants(_Select_Selection.Marque);

                mdl.eSelectMulti(listeComps, _Select_Selection.Marque, false);

                mdl.eAssemblyDoc().Isolate();

                mdl.eSelectMulti(listeCorps, _Select_Selection.Marque, false);

                Isoler = true;

                _Button_IsolerComposants.Caption = "Afficher tout les composants";
            }
            else
            {
                ExitIsoler();
            }
        }

        private void SelectionChanged(Object sender, int Item)
        {
            var TextListBox = sender as CtrlTextListBox;

            var mdl = App.ModelDoc2;

            var listeElements = Bdd.ListeFace(TextListBox.SelectedText);

            var listeFace = new List<Face2>();

            foreach (var element in listeElements)
            {
                if (element.Key.IsHidden(false) && !ListeHiddenComposants.Contains(element.Key.eKeyAvecConfig()))
                {
                    element.Key.eSelectById(mdl, 16, true);
                    element.Key.eDeSelectById(mdl);
                    ListeHiddenComposants.Add(element.Key.eKeyAvecConfig());
                }

                foreach (var face in element.Value)
                {
                    if (face.IsNull()) continue;

                    listeFace.Add(face);
                }
            }

            ExitIsoler();

            mdl.eSelectMulti(listeFace, _Select_Selection.Marque, false);
        }

        private void AfficherPercage()
        {
            Rechercher_Percage();
            _TextListBox_Percage.Liste = Bdd.ListeDiametres();
        }

        private BDD Bdd;

        private void Rechercher_Percage()
        {
            Bdd = new BDD();

            var mdl = App.ModelDoc2;

            App.ModelDoc2.eRecParcourirComposants(
                    c =>
                    {
                        if (!c.IsHidden(true) && (c.TypeDoc() == eTypeDoc.Piece))
                        {
                            Bdd.Decompter(c);
                        }

                        return false;
                    }
                );
        }

        public class BDD
        {
            private Dictionary<Double, Dictionary<Component2, List<Face2>>> Dic = new Dictionary<Double, Dictionary<Component2, List<Face2>>>();

            public Boolean Decompter(Component2 cp)
            {
                try
                {
                    if (!cp.IsHidden(true))
                    {
                        foreach (var corps in cp.eListeCorps())
                        {
                            foreach (var face in corps.eListeDesFaces())
                            {
                                Surface S = face.GetSurface();
                                if (S.IsRef() && S.IsCylinder() && (face.GetLoopCount() > 1))
                                {
                                    Double[] ListeParam = (Double[])S.CylinderParams;
                                    Double Diam = Math.Round(ListeParam[6] * 2.0 * 1000, 2);

                                    var dicDiam = new Dictionary<Component2, List<Face2>>();
                                    if (Dic.ContainsKey(Diam))
                                        dicDiam = Dic[Diam];
                                    else
                                        Dic.Add(Diam, dicDiam);

                                    var listFace = new List<Face2>();
                                    if (dicDiam.ContainsKey(cp))
                                        listFace = dicDiam[cp];
                                    else
                                        dicDiam.Add(cp, listFace);

                                    listFace.Add(face);
                                }
                            }
                        }
                    }

                }
                catch (Exception e) { this.LogMethode(new Object[] { e }); }

                return false;
            }

            private Dictionary<String, Double> _DicIntitule = new Dictionary<String, Double>();

            public List<String> ListeDiametres()
            {
                int lgInt = 0;
                int lgNb = 0;
                foreach (var Diam in Dic.Keys)
                {
                    int nb = 0;
                    foreach (var list in Dic[Diam].Values)
                        nb += list.Count;

                    lgInt = Math.Max(lgInt, Diam.ToString().Length);
                    lgNb = Math.Max(lgNb, nb.ToString().Length);
                }

                String format = "Ø {0,-" + lgInt + "}  × {1,-" + lgNb + "}";

                foreach (var Diam in Dic.Keys)
                {
                    int nb = 0;
                    foreach (var list in Dic[Diam].Values)
                        nb += list.Count;

                    _DicIntitule.Add(String.Format(format, Diam, nb), Diam);
                }

                return _DicIntitule.Keys.ToList();
            }

            public Dictionary<Component2, List<Face2>> ListeFace(String Intitule)
            {
                return Dic[_DicIntitule[Intitule]];
            }
        }

        protected void RunOkCommand()
        {
            var mdl = App.ModelDoc2;

            List<Body2> listeCorps = mdl.eSelect_RecupererListeObjets<Body2>(_Select_Selection.Marque);

            mdl.eSelectMulti(listeCorps, -1, false);
        }
    }
}
