using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace ModuleProduction.ModuleControlerRepere
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Controler les repères"),
        ModuleNom("ControlerRepere"),
        ModuleDescription("Controler les repères des corps.")
        ]
    public class PageControlerRepere : BoutonPMPManager
    {
        private const String Erreur = "Aucun";

        private ModelDoc2 MdlBase;

        public PageControlerRepere()
        {
            try
            {
                InitModeleBase();

                OnCalque += Calque;
                OnRunAfterActivation += AfficherReperes;
                OnRunOkCommand += RunOkCommand;
                OnRunAfterClose += ExitIsoler;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private CtrlTextListBox _TextListBox_Reperes;
        private CtrlCheckBox _CheckBox_ComposantsCache;
        private CtrlTextBox _TextBox_Nb;
        private CtrlSelectionBox _Select_Selection;

        private CtrlButton _Button_IsolerComposants;

        protected void Calque()
        {
            try
            {
                Groupe G;

                if (MdlBase.TypeDoc() == eTypeDoc.Assemblage)
                {
                    G = _Calque.AjouterGroupe("Options");

                    _Button_IsolerComposants = G.AjouterBouton("Isoler les composants");
                    _Button_IsolerComposants.OnButtonPress += delegate (Object sender) { IsolerComposants(); };
                }

                G = _Calque.AjouterGroupe("Repères :");

                _TextListBox_Reperes = G.AjouterTextListBox();
                _TextListBox_Reperes.TouteHauteur = true;
                _TextListBox_Reperes.Height = 170;
                _TextListBox_Reperes.SelectionMultiple = false;
                _TextListBox_Reperes.OnSelectionChanged += SelectionChanged;

                _CheckBox_ComposantsCache = G.AjouterCheckBox("Prendre en compte les composants cachés");
                _CheckBox_ComposantsCache.OnIsCheck += delegate (Object sender, Boolean value) { AfficherReperes(); };

                G = _Calque.AjouterGroupe("Selection");

                _TextBox_Nb = G.AjouterTexteBox("Nb de corps :");
                _TextBox_Nb.LectureSeule = true;

                _Select_Selection = G.AjouterSelectionBox("");
                _Select_Selection.SelectionMultipleMemeEntite = true;
                _Select_Selection.SelectionDansMultipleBox = true;
                _Select_Selection.UneSeuleEntite = false;
                _Select_Selection.FiltreSelection(swSelectType_e.swSelSOLIDBODIES, swSelectType_e.swSelCOMPONENTS);
                _Select_Selection.Hauteur = 15;
                _Select_Selection.OnSelectionChanged += delegate (Object sender, int nb) { _TextBox_Nb.Text = nb.ToString(); };
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private HashSet<String> ListeHiddenComposants = new HashSet<string>();

        private Boolean Isoler = false;

        private void InitModeleBase()
        {
            MdlBase = App.ModelDoc2;
        }

        private void ExitIsoler()
        {
            if (MdlBase.TypeDoc() != eTypeDoc.Assemblage)
                return;

            MdlBase.eAssemblyDoc().ExitIsolate();
            _Button_IsolerComposants.Caption = "Isoler les composants";
            Isoler = false;
        }

        private void IsolerComposants()
        {
            if (MdlBase.TypeDoc() != eTypeDoc.Assemblage)
                return;

            if (!Isoler)
            {
                List<Body2> listeCorps = MdlBase.eSelect_RecupererListeObjets<Body2>(_Select_Selection.Marque);
                List<Component2> listeComps = MdlBase.eSelect_RecupererListeComposants(_Select_Selection.Marque);

                MdlBase.eSelectMulti(listeComps, _Select_Selection.Marque, false);

                MdlBase.eAssemblyDoc().Isolate();

                MdlBase.eSelectMulti(listeCorps, _Select_Selection.Marque, false);

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
            var listeElements = Bdd.ListeCorps(Item);

            var listeCorps = new List<Body2>();

            foreach (var element in listeElements)
            {
                if (element.Composant.IsHidden(false) && !ListeHiddenComposants.Contains(element.Composant.eKeyAvecConfig()))
                {
                    element.Composant.eSelectById(MdlBase, 16, true);
                    element.Composant.eDeSelectById(MdlBase);
                    ListeHiddenComposants.Add(element.Composant.eKeyAvecConfig());
                }

                //WindowLog.EcrireF("{0} {1}", element.Composant.Name2, element.Corps.Name);
                var corps = element.Composant.eChercherCorps(element.Corps.Name, false);
                if (corps.IsNull()) continue;

                listeCorps.Add(corps);
            }

            ExitIsoler();

            MdlBase.eSelectMulti(listeCorps, _Select_Selection.Marque, false);
        }

        private void AfficherReperes()
        {
            Rechercher_Reperes();
            _TextListBox_Reperes.Liste = Bdd.ListeNoms();
        }

        private BDD Bdd;

        private void Rechercher_Reperes()
        {
            Bdd = new BDD();

            Predicate<Component2> Test = c =>
            {
                if (c.ExcludeFromBOM)
                    return false;

                bool filtre = false;

                if (_CheckBox_ComposantsCache.IsChecked)
                    filtre = c.IsSuppressed();
                else
                    filtre = c.IsHidden(true);

                if (!filtre && (c.TypeDoc() == eTypeDoc.Piece))
                {

                    ModelDoc2 mdl = c.eModelDoc2();
                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                    mdl.ShowConfiguration2(c.eNomConfiguration());
                    mdl.EditRebuild3();
                    PartDoc piece = mdl.ePartDoc();

                    foreach (var dossier in piece.eListeDesDossiersDePiecesSoudees(d => { return !d.eEstExclu(); }))
                        Bdd.AjouterDossier(dossier, c);
                }

                return false;
            };

            if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                Test(MdlBase.eComposantRacine());
            else
                MdlBase.eRecParcourirComposants(Test, c => { if (c.ExcludeFromBOM) return false; return true; });

            MdlBase.eActiver();
        }

        public class BDD
        {
            private Dictionary<String, List<Element>> Dic = new Dictionary<String, List<Element>>();

            public void AjouterDossier(BodyFolder dossier, Component2 comp)
            {
                Body2 corps = dossier.ePremierCorps();

                if (corps.IsNull()) return;

                String Repere = "";

                if (dossier.ePropExiste(CONSTANTES.REF_DOSSIER))
                    Repere = dossier.eProp(CONSTANTES.REF_DOSSIER);
                else
                    Repere = Erreur;

                Ajouter(Repere, dossier, comp);
            }

            private void Ajouter(String repere, BodyFolder dossier, Component2 cp)
            {
                List<Element> listElements = new List<Element>();
                if (Dic.ContainsKey(repere))
                    listElements = Dic[repere];
                else
                    Dic.Add(repere, listElements);

                foreach (var corps in dossier.eListeDesCorps())
                    listElements.Add(new Element(corps, cp));
            }

            private List<String> _ListNoms = null;

            public List<String> ListeNoms()
            {
                if (_ListNoms.IsNull())
                {
                    _ListNoms = new List<String>();
                    foreach (var r in Dic)
                        _ListNoms.Add(String.Format("{0,-7}{1,5}", r.Key, "×" + r.Value.Count));

                    _ListNoms.Sort(new WindowsStringComparer());
                }

                return _ListNoms;
            }

            public List<Element> ListeCorps(int Niveau)
            {
                List<Element> liste = new List<Element>();

                String rep = _ListNoms[Niveau].Split(new char[] { ' ' })[0].Trim();

                return Dic[rep];
            }

            public class Element
            {
                public Component2 Composant { get; private set; }
                public Body2 Corps { get; private set; }
                public Object Pid { get; private set; }

                public Element(Body2 corps, Component2 composant)
                {
                    Composant = composant;
                    Corps = corps;
                    Pid = Composant.eModelDoc2().Extension.GetPersistReference3(Corps);
                }
            }
        }

        protected void RunOkCommand()
        {
            List<Body2> listeCorps = MdlBase.eSelect_RecupererListeObjets<Body2>(_Select_Selection.Marque);

            MdlBase.eSelectMulti(listeCorps, -1, false);
        }
    }
}
