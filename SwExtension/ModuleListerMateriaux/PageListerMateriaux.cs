using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace ModuleListerMateriaux
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Lister les materiaux et les profils"),
        ModuleNom("ListerMateriauxEtProfils"),
        ModuleDescription("Lister les materiaux et les profils.")
        ]
    public class PageListerMateriaux : BoutonPMPManager
    {
        private const String NomVolume = "Volume";

        public PageListerMateriaux()
        {
            try
            {
                OnCalque += Calque;
                OnRunAfterActivation += AfficherMateriaux;
                OnRunOkCommand += RunOkCommand;
                OnRunAfterClose += ExitIsoler;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private CtrlTextListBox _TextListBox_Materiaux;
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

                G = _Calque.AjouterGroupe("Materiaux :");

                _TextListBox_Materiaux = G.AjouterTextListBox();
                _TextListBox_Materiaux.TouteHauteur = true;
                _TextListBox_Materiaux.Height = 170;
                _TextListBox_Materiaux.SelectionMultiple = false;
                _TextListBox_Materiaux.OnSelectionChanged += SelectionChanged;

                _CheckBox_ComposantsCache = G.AjouterCheckBox("Prendre en compte les composants cachés");
                _CheckBox_ComposantsCache.OnIsCheck += delegate (Object sender, Boolean value) { AfficherMateriaux(); };

                G = _Calque.AjouterGroupe("Selection");

                _TextBox_Nb = G.AjouterTexteBox("Nb de corps :");
                _TextBox_Nb.LectureSeule = true;

                _Select_Selection = G.AjouterSelectionBox("");
                _Select_Selection.SelectionMultipleMemeEntite = true;
                _Select_Selection.SelectionDansMultipleBox = true;
                _Select_Selection.UneSeuleEntite = false;
                _Select_Selection.FiltreSelection(swSelectType_e.swSelSOLIDBODIES , swSelectType_e.swSelCOMPONENTS);
                _Select_Selection.Hauteur = 15;
                _Select_Selection.OnSelectionChanged += delegate (Object sender, int nb) { _TextBox_Nb.Text = nb.ToString(); };
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private HashSet<String> ListeHiddenComposants = new HashSet<string>();

        private Boolean Isoler = false;

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

                var corps = element.GetCorps();
                if (corps.IsNull()) continue;

                listeCorps.Add(corps);
            }

            ExitIsoler();

            MdlBase.eSelectMulti(listeCorps, _Select_Selection.Marque, false);
        }

        private void AfficherMateriaux()
        {
            Rechercher_Materiaux();
            _TextListBox_Materiaux.Liste = Bdd.ListeNoms();
        }

        private BDD Bdd;

        private void Rechercher_Materiaux()
        {
            try
            {
                Bdd = new BDD();

                Predicate<Component2> Test = c =>
                {
                    if (c.TypeDoc() == eTypeDoc.Piece && (!c.IsSuppressed()))
                    {
                        bool filtre = true;

                        if (!_CheckBox_ComposantsCache.IsChecked)
                            filtre = c.Visible == (int)swComponentVisibilityState_e.swComponentVisible;

                        if (filtre)
                        {
                            foreach (var dossier in c.eListeDesDossiersDePiecesSoudees())
                                Bdd.AjouterDossier(dossier, c);
                        }
                    }

                    return false;
                };

                if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                    Test(MdlBase.eComposantRacine());
                else
                    MdlBase.eRecParcourirComposants(Test);
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        public class BDD
        {
            private SortedDictionary<String, SortedDictionary<String, SortedDictionary<String, List<Element>>>> Dic = new SortedDictionary<String, SortedDictionary<String, SortedDictionary<String, List<Element>>>>(new WindowsStringComparer());

            public void AjouterDossier(BodyFolder dossier, Component2 comp)
            {
                Body2 corps = dossier.ePremierCorps();

                if (corps.IsNull()) return;

                String BaseMateriau;
                String Materiau = corps.eGetMateriauCorpsOuComp(comp, out BaseMateriau);
                string Profil;
                if (corps.eTypeDeCorps() == eTypeCorps.Tole)
                {
                    Double Ep = corps.eEpaisseurCorpsOuDossier(dossier);
                    if (Ep == -1)
                    {
                        WindowLog.EcrireF("Pb d'epaisseur sur le corps {0}", corps.Name);
                        return;
                    }

                    Profil = "Ep " + Ep.ToString();
                }
                else if (dossier.ePropExiste(CONSTANTES.PROFIL_NOM))
                    Profil = dossier.eProp(CONSTANTES.PROFIL_NOM);
                else
                    Profil = NomVolume;

                Ajouter(BaseMateriau, Materiau, Profil, dossier, comp);
            }

            private void Ajouter(String baseMateriau, String materiau, String profil, BodyFolder dossier, Component2 cp)
            {
                SortedDictionary<String, SortedDictionary<String, List<Element>>> dicMat = new SortedDictionary<String, SortedDictionary<String, List<Element>>>(new WindowsStringComparer());
                if (Dic.ContainsKey(baseMateriau))
                    dicMat = Dic[baseMateriau];
                else
                    Dic.Add(baseMateriau, dicMat);

                SortedDictionary<String, List<Element>> dicProfil = new SortedDictionary<String, List<Element>>(new WindowsStringComparer());
                if (dicMat.ContainsKey(materiau))
                    dicProfil = dicMat[materiau];
                else
                    dicMat.Add(materiau, dicProfil);

                List<Element> listElements = new List<Element>();
                if (dicProfil.ContainsKey(profil))
                    listElements = dicProfil[profil];
                else
                    dicProfil.Add(profil, listElements);

                foreach (var corps in dossier.eListeDesCorps())
                    listElements.Add(new Element(corps, cp));
            }

            private List<int[]> ListIndex = new List<int[]>();

            public List<String> ListeNoms()
            {
                var liste = new List<String>();
                int i = 0;
                foreach (var BaseMateriau in Dic.Keys)
                {
                    liste.Add(BaseMateriau);
                    ListIndex.Add(new int[] { i, -1, -1 });
                    var listeMateriau = Dic[BaseMateriau];
                    int j = 0;
                    foreach (var Materiau in listeMateriau.Keys)
                    {
                        liste.Add("  " + Materiau);
                        ListIndex.Add(new int[] { i, j, -1 });
                        var listeProfil = listeMateriau[Materiau];
                        int k = 0;
                        foreach (var Profil in listeProfil.Keys)
                        {
                            liste.Add("    " + Profil);
                            ListIndex.Add(new int[] { i, j, k });
                            k++;
                        }
                        j++;
                    }
                    i++;
                }

                return liste;
            }

            public List<Element> ListeCorps(int Niveau)
            {
                List<Element> liste = new List<Element>();

                int[] n = ListIndex[Niveau];

                int i = 0;
                foreach (var BaseMateriau in Dic.Keys)
                {
                    if (n[0] == i)
                    {
                        int j = 0;
                        var listeMateriau = Dic[BaseMateriau];
                        foreach (var Materiau in listeMateriau.Keys)
                        {
                            if ((n[1] == -1) || (n[1] == j))
                            {
                                int k = 0;
                                var listeProfil = listeMateriau[Materiau];
                                foreach (var Profil in listeProfil.Keys)
                                {
                                    if ((n[2] == -1) || (n[2] == k))
                                    {
                                        var listeElements = listeProfil[Profil];
                                        liste.AddRange(listeElements);
                                    }
                                    k++;
                                }
                            }
                            j++;
                        }
                    }
                    i++;
                }

                return liste;
            }

            public class Element
            {
                public Component2 Composant { get; private set; }
                public Body2 Corps { get; private set; }

                public Element(Body2 corps, Component2 composant)
                {
                    Composant = composant;
                    Corps = corps;
                }

                public Body2 GetCorps()
                {
                    return Composant.eChercherCorps(Corps.Name, false);
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
