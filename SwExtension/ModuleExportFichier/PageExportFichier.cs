using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModuleExportFichier
{
    public static class SwAttributDernierDossier
    {
        private const String ATTRIBUT_NOM = "DernierDossier";
        private const String ATTRIBUT_PARAM = "Chemin";

        private static AttributeDef AttDef = null;

        private static Parameter eAttributDernierDossier(this DrawingDoc dessin)
        {
            if (AttDef.IsNull())
            {
                AttDef = App.Sw.DefineAttribute(ATTRIBUT_NOM);
                AttDef.AddParameter(ATTRIBUT_PARAM, (int)swParamType_e.swParamTypeString, 0, 0);
                AttDef.Register();
            }

            // Recherche de l'attribut dans la piece
            SolidWorks.Interop.sldworks.Attribute Att = null;
            ModelDoc2 mdl = dessin.eModelDoc2();
            Parameter P = null;
            Feature F = mdl.eChercherFonction(f => { return f.Name == ATTRIBUT_NOM; });
            if (F.IsRef())
            {
                Att = F.GetSpecificFeature2();

                P = (Parameter)Att.GetParameter(ATTRIBUT_PARAM);

                if (P.IsNull())
                {
                    Att.Delete(false);
                    Att = null;
                }
            }

            if (Att.IsNull())
            {
                Att = AttDef.CreateInstance5(mdl, null, ATTRIBUT_NOM, 1, (int)swInConfigurationOpts_e.swAllConfiguration);
                P = (Parameter)Att.GetParameter(ATTRIBUT_PARAM);
            }

            return P;
        }

        public static void eSetDernierDossier(this DrawingDoc dessin, String val)
        {
            Parameter P = dessin.eAttributDernierDossier();
            P.SetStringValue2(val, (int)swInConfigurationOpts_e.swAllConfiguration, "");
        }

        public static String eGetDernierDossier(this DrawingDoc dessin)
        {
            Parameter P = dessin.eAttributDernierDossier();
            return P.GetStringValue();
        }
    }

    public static class OutilsCommun
    {
        public static String CheminRelatif(String dossier, String chemin)
        {
            Uri u1 = new Uri(PathAddBackslash(dossier));
            Uri u2 = new Uri(chemin);

            return Uri.UnescapeDataString(u1.MakeRelativeUri(u2).OriginalString).Replace('/', '\\');
        }

        public static string PathAddBackslash(string path)
        {
            string separator1 = Path.DirectorySeparatorChar.ToString();
            string separator2 = Path.AltDirectorySeparatorChar.ToString();
            path = path.TrimEnd();

            if (path.EndsWith(separator1) || path.EndsWith(separator2))
                return path;

            if (path.Contains(separator2))
                return path + separator2;

            return path + separator1;
        }
    }

    public abstract class PageExportFichier : BoutonPMPManager
    {
        protected Parametre FormatExport;
        protected Parametre GroupeDernierDossier;
        protected Parametre GroupeSelectionnerDossier;
        protected Parametre ToutesLesFeuilles;

        protected CtrlCheckBox _CheckBox_ToutesLesFeuilles;

        public PageExportFichier()
        {
            GroupeDernierDossier = _Config.AjouterParam("DernierDossier", false, "Dernier dossier");
            GroupeSelectionnerDossier = _Config.AjouterParam("SelectionnerDossier", true, "Selectionner un dossier");
            ToutesLesFeuilles = _Config.AjouterParam("ToutesLesFeuilles", false, "Toutes les feuilles");

            OnRunAfterClose += RecupererInfosFichier;
        }

        protected Groupe G;

        protected Dossier _DernierDossier;
        protected Dossier _SelectionnerDossier;

        public void AjouterCalqueDossier()
        {
            _CheckBox_ToutesLesFeuilles = G.AjouterCheckBox(ToutesLesFeuilles);

            String CheminDernierDossier = App.DrawingDoc.eGetDernierDossier();

            _DernierDossier = new Dossier(_Calque, GroupeDernierDossier, CheminDernierDossier, App.ModelDoc2.eNomSansExt(), FormatExport.GetValeur<eTypeFichierExport>());
            _SelectionnerDossier = new Dossier(_Calque, GroupeSelectionnerDossier, OutilsCommun.CheminRelatif(App.ModelDoc2.eDossier(), App.ModelDoc2.eDossier()), App.ModelDoc2.eNomSansExt(), FormatExport.GetValeur<eTypeFichierExport>(), true, true);

            _CheckBox_ToutesLesFeuilles.OnIsCheck += delegate (Object sender, Boolean value)
            {
                    String n = ((Sheet)App.DrawingDoc.GetCurrentSheet()).GetName();
                    if (value)
                        n = App.ModelDoc2.eNomSansExt();

                    _DernierDossier.NomFichierOriginal = n;
                    _SelectionnerDossier.NomFichierOriginal = n;
            };

            _DernierDossier.Groupe.OnIsCheck += delegate (Object sender, Boolean value)
            {
                if (value)
                    _SelectionnerDossier.Groupe.UnCheck(null);
                else
                    _SelectionnerDossier.Groupe.Check(null);
            };

            _SelectionnerDossier.Groupe.OnIsCheck += delegate (Object sender, Boolean value)
            {
                if (value)
                    _DernierDossier.Groupe.UnCheck(null);
                else
                    _DernierDossier.Groupe.Check(null);
            };

            OnRunAfterActivation += _DernierDossier.Maj;
            OnRunAfterActivation += _SelectionnerDossier.Maj;

            if (String.IsNullOrWhiteSpace(CheminDernierDossier) || !Directory.Exists(Path.Combine(App.ModelDoc2.eDossier(), CheminDernierDossier)))
            {
                _DernierDossier.Groupe.IsChecked = false;
                _DernierDossier.Groupe.Visible = false;

                _SelectionnerDossier.Groupe.IsChecked = true;
                _SelectionnerDossier.Groupe.Expanded = true;
                _SelectionnerDossier.Groupe.OnUnCheck += _SelectionnerDossier.Groupe.Check;

                OnRunAfterActivation -= _DernierDossier.Maj;
            }
            else
            {
                _DernierDossier.Groupe.IsChecked = true;
                _DernierDossier.Groupe.Expanded = true;

                _SelectionnerDossier.Groupe.IsChecked = false;
                _SelectionnerDossier.Groupe.Expanded = false;
            }

            _CheckBox_ToutesLesFeuilles.ApplyParam();
        }

        protected String NomDossier;
        protected String NomFichier;
        protected String NomFichierComplet;
        protected String Indice;

        protected void RecupererInfosFichier()
        {
            if (_SelectionnerDossier.Groupe.IsChecked)
            {
                NomDossier = _SelectionnerDossier.NomDossierComplet;
                NomFichier = _SelectionnerDossier.NomFichier;
                NomFichierComplet = _SelectionnerDossier.NomFichierComplet;
                Indice = _SelectionnerDossier.Indice;

                App.DrawingDoc.eSetDernierDossier(_SelectionnerDossier.NomDossierRelatifSvg);
            }
            else if (_DernierDossier.Groupe.IsChecked)
            {
                NomDossier = _DernierDossier.NomDossierComplet;
                NomFichier = _DernierDossier.NomFichier;
                NomFichierComplet = _DernierDossier.NomFichierComplet;
                Indice = _DernierDossier.Indice;
            }
        }

        protected class Dossier
        {
            private Calque _Calque;

            private GroupeAvecCheckBox _Groupe;

            public GroupeAvecCheckBox Groupe { get { return _Groupe; } }

            private CtrlTextBox _TextBox_CheminDossier;
            private CtrlButton _Button_Parcourir;

            private CtrlTextListBox _TextListBox_ListeDossiers;

            private CtrlCheckBox _CheckBox_CreerDossier;
            private CtrlTextBox _TextBox_NomNvxDossier;
            private CtrlButton _Button_CreerDossier;

            private CtrlTextListBox _TextListBox_ListeFichiers;
            private CtrlTextBox _TextBox_NomNvxFichier;

            private Parametre _ParamGroupe;
            private String _NomDossierRelatif;
            private String _NomDossierCourant = DossierCourant;
            private String _NomDossierComplet;
            private String _NomFichierOriginal;
            private String _NomFichierBase;
            private String _NomFichierComplet;
            private String _Indice;

            private Boolean _Selectionnable = false;
            private Boolean _AjouterIndiceDossier = false;

            private eTypeFichierExport _TypeFichier;

            public eTypeFichierExport TypeFichier { get { return _TypeFichier; } set { _TypeFichier = value; } }


            public String NomDossierComplet { get { CheminDossierComplet(); return _NomDossierComplet; } }
            public String NomDossierRelatifSvg
            {
                get
                {
                    if (_NomDossierCourant == DossierCourant)
                        return _NomDossierRelatif;

                    return Path.Combine(_NomDossierRelatif, _NomDossierCourant);
                }
            }
            public String NomFichierOriginal { get { return _NomFichierOriginal; } set { _NomFichierOriginal = value; Maj(); } }
            public String NomFichier { get { return _NomFichierBase; } set { _NomFichierBase = value; Maj(); } }
            public String NomFichierComplet { get { return Path.GetFileNameWithoutExtension(_NomFichierComplet); } }
            public String Indice { get { return _Indice; } }

            protected const String DossierCourant = ".";

            public Dossier(Calque Calque, Parametre paramGroupe, String dossier, String fichier, eTypeFichierExport typeFichier, Boolean selectionnable = false, Boolean ajouterIndiceDossier = false)
            {
                _Calque = Calque;
                _ParamGroupe = paramGroupe;
                _NomDossierRelatif = dossier;
                _NomFichierBase = fichier;
                _NomFichierOriginal = fichier;

                _TypeFichier = typeFichier;
                _Selectionnable = selectionnable;
                _AjouterIndiceDossier = ajouterIndiceDossier;

                AjouterAuCalque();
            }

            public void AjouterAuCalque()
            {

                _Groupe = _Calque.AjouterGroupeAvecCheckBox(_ParamGroupe);

                _TextBox_CheminDossier = _Groupe.AjouterTexteBox();
                AfficherCheminDossier();
                _TextBox_CheminDossier.LectureSeule = true;

                if (_Selectionnable)
                {
                    _Button_Parcourir = _Groupe.AjouterBouton("Parcourir");
                    _Button_Parcourir.OnButtonPress += delegate (Object Bouton)
                    {
                        System.Windows.Forms.FolderBrowserDialog pDialogue = new System.Windows.Forms.FolderBrowserDialog();
                        pDialogue.SelectedPath = App.ModelDoc2.eDossier();

                        if (pDialogue.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            _NomDossierRelatif = OutilsCommun.CheminRelatif(App.ModelDoc2.eDossier(), pDialogue.SelectedPath);
                            AfficherCheminDossier();
                        }

                        Maj();
                    };

                    _TextListBox_ListeDossiers = _Groupe.AjouterTextListBox("Liste des dossiers");
                    _TextListBox_ListeDossiers.TouteHauteur = true;
                    _TextListBox_ListeDossiers.Height = 50;

                    _TextListBox_ListeDossiers.OnSelectionChanged += delegate (Object sender, int Item)
                    {
                        _NomDossierCourant = _TextListBox_ListeDossiers.SelectedText;
                        MajListeFichiers();
                    };

                    _CheckBox_CreerDossier = _Groupe.AjouterCheckBox("Nouveau dossier");
                    _TextBox_NomNvxDossier = _Groupe.AjouterTexteBox();
                    _Button_CreerDossier = _Groupe.AjouterBouton("Creer");
                    _CheckBox_CreerDossier.OnIsCheck += _TextBox_NomNvxDossier.IsEnable;
                    _CheckBox_CreerDossier.OnIsCheck += _Button_CreerDossier.IsEnable;
                    _CheckBox_CreerDossier.OnIsCheck += _TextBox_NomNvxDossier.IsVisible;
                    _CheckBox_CreerDossier.OnIsCheck += _Button_CreerDossier.IsVisible;
                    _CheckBox_CreerDossier.IsChecked = false;

                    _Button_CreerDossier.OnButtonPress += delegate (Object Bouton)
                    {
                        String n = _TextBox_NomNvxDossier.Text;
                        Directory.CreateDirectory(Path.Combine(CheminDossierRelatif(), n));

                        Maj();

                        _CheckBox_CreerDossier.IsChecked = false;
                        _TextListBox_ListeDossiers.SelectedText = n;
                    };
                }

                _TextListBox_ListeFichiers = _Groupe.AjouterTextListBox("Contenu du dossier :");
                _TextListBox_ListeFichiers.TouteHauteur = true;
                _TextListBox_ListeFichiers.Height = 50;

                _TextBox_NomNvxFichier = _Groupe.AjouterTexteBox();
                _TextBox_NomNvxFichier.OnTextBoxChanged += delegate (Object sender, String text)
                {
                    if (!EcraserInfos) return;

                    _NomFichierComplet = text;
                };
            }

            private Boolean EcraserInfos = true;

            private void AfficherCheminDossier()
            {
                String texte = _NomDossierRelatif;
                if (texte.StartsWith(@"..\"))
                    texte = texte.Remove(0, 3);
                else
                    texte = @"\" + texte;

                _TextBox_CheminDossier.Text = texte;
            }

            private String CheminDossierRelatif()
            {
                String DossierParent = App.ModelDoc2.eDossier();

                String n = Path.GetFullPath(Path.Combine(DossierParent, _NomDossierRelatif));

                if (!Directory.Exists(n))
                {
                    _NomDossierRelatif = DossierCourant;
                    n = DossierParent;
                }

                return n;
            }

            private void CheminDossierComplet()
            {
                _NomDossierComplet = Path.GetFullPath(Path.Combine(CheminDossierRelatif(), _NomDossierCourant)); ;
            }

            public void Maj()
            {
                try
                {
                    CheminDossierComplet();

                    if (_Selectionnable)
                        MajListeDossiers();

                    MajListeFichiers();
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private void MajListeDossiers()
            {
                if (_Selectionnable)
                    _TextListBox_ListeDossiers.Liste = ListeDossiers();
            }

            private void MajListeFichiers()
            {
                try
                {
                    CheminDossierComplet();

                    EcraserInfos = false;
                    _TextBox_NomNvxFichier.Text = NomNvxFichier();
                    EcraserInfos = true;

                    if (_Selectionnable)
                        _TextBox_NomNvxDossier.Text = NomNvxDossier();

                    _TextListBox_ListeFichiers.Liste = ListeFichiers(new DirectoryInfo(_NomDossierComplet));
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private List<String> ListeDossiers()
            {
                List<String> l = new List<string>();

                DirectoryInfo D = new DirectoryInfo(CheminDossierRelatif());

                foreach (var d in D.EnumerateDirectories())
                {
                    l.Add(d.Name);
                }

                l.Insert(0, DossierCourant);

                return l;
            }

            private readonly String ChaineIndice = "ZYXWVUTSRQPONMLKJIHGFEDCBA";

            private String NomNvxDossier()
            {
                String NomDossier = _TypeFichier.GetEnumInfo<Intitule>().ToUpperInvariant();

                if (!_AjouterIndiceDossier)
                    return NomDossier;

                DirectoryInfo D = new DirectoryInfo(CheminDossierRelatif());
                List<String> ListeD = new List<string>();

                foreach (var d in D.GetDirectories())
                {
                    if (d.Name.ToUpperInvariant().StartsWith(NomDossier))
                    {
                        ListeD.Add(d.Name);
                    }
                }

                ListeD.Sort(new WindowsStringComparer(ListSortDirection.Ascending));

                String i = ChercherIndice(ListeD);

                return NomDossier + " - " + i;
            }

            private String NomNvxFichier()
            {
                ChercherIndiceFichier();

                _NomFichierBase = _NomFichierOriginal;

                _NomFichierComplet = _NomFichierBase + " - " + _Indice + TypeFichier.GetEnumInfo<ExtFichier>();

                return _NomFichierComplet;
            }

            private List<String> ListeFichiers(DirectoryInfo D, Boolean avecExtension = true)
            {
                List<String> ListeF = new List<string>();

                foreach (var f in D.GetFiles())
                {
                    if (f.Extension.ToLowerInvariant() == _TypeFichier.GetEnumInfo<ExtFichier>().ToLowerInvariant())
                        ListeF.Add(avecExtension ? f.Name : Path.GetFileNameWithoutExtension(f.Name));
                }

                ListeF.Sort(new WindowsStringComparer(ListSortDirection.Ascending));

                return ListeF;
            }

            private void ChercherIndiceFichier()
            {

                DirectoryInfo D = new DirectoryInfo(_NomDossierComplet);
                _Indice = ChercherIndiceDossierParent(D.Name);

                if (!String.IsNullOrWhiteSpace(_Indice))
                    return;

                List<String> ListeF = ListeFichiers(D, false);

                _Indice = ChercherIndice(ListeF);
            }

            private String ChercherIndiceDossierParent(String s)
            {
                String r = "";
                foreach (var i in ChaineIndice)
                {
                    r = " Ind " + i;

                    if (s.EndsWith(r))
                        return r.Trim();
                }

                return "";
            }

            private String ChercherIndice(List<String> liste)
            {
                for (int i = 0; i < ChaineIndice.Length; i++)
                {
                    if (liste.Any(d => { return d.EndsWith(" Ind " + ChaineIndice[i]) ? true : false; }))
                        return "Ind " + ChaineIndice[Math.Max(0, i - 1)];
                }

                return "Ind " + ChaineIndice.Last();
            }

        }
    }
}
