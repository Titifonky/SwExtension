using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SwExtension
{
    [ModuleTitre("Gestion des parametres"), ModuleNom("OngletParametres")]
    public partial class OngletParametres : Form
    {

        private SldWorks _Sw = null;

        public int b = 5;
        public int j = 3;

        private ContextMenuStrip ParamContextMenuListBox;

        private ContextMenuStrip PropContextMenuListBox;

        private HashSet<String> HashPropExclu = new HashSet<String>();

        public OngletParametres(SldWorks sw)
        {
            _Sw = sw;

            InitializeComponent();

            _DicParam = new DicParam(ListBoxParams);
            _DicPropModele = new DicProp(ListBoxPropMdl);
            _DicPropCfg = new DicProp(ListBoxPropCfg);

            ToolStripMenuItem ParamMenuItem = new ToolStripMenuItem { Text = "Supprimer" };
            ParamMenuItem.Click += SupprimerParam;

            ParamContextMenuListBox = new ContextMenuStrip();
            ParamContextMenuListBox.Items.AddRange(new ToolStripItem[] { ParamMenuItem });
            ListBoxParams.MouseDown += ListBoxParams_MouseDown;
            ListBoxParams.ContextMenuStrip = ParamContextMenuListBox;

            ToolStripMenuItem PropMenuItem = new ToolStripMenuItem { Text = "Supprimer" };
            PropMenuItem.Click += SupprimerProp;

            PropContextMenuListBox = new ContextMenuStrip();
            PropContextMenuListBox.Items.AddRange(new ToolStripItem[] { PropMenuItem });
            ListBoxPropMdl.MouseDown += ListBoxProp_MouseDown;
            ListBoxPropMdl.ContextMenuStrip = PropContextMenuListBox;
            ListBoxPropCfg.MouseDown += ListBoxProp_MouseDown;
            ListBoxPropCfg.ContextMenuStrip = PropContextMenuListBox;

            String NomModule = GetType().GetModuleNom();
            String TitreModule = GetType().GetModuleTitre();
            ConfigModule _Config = new ConfigModule(NomModule, TitreModule);

            Parametre pListePropExclu = _Config.AjouterParam("ListePropExclu", "Adresse,Client,Description,Designation,Dessinateur,Empreinte,ExcluNomenclature,LierConfigs,MasseModele,Matériau,NoClient,NoCommande,PrefixeEmpreinte,Quantite,TarauderEmpreinte", "Liste des proprietes à exclure");

            HashPropExclu = new HashSet<string>(pListePropExclu.GetValeur<String>().Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries));

            ResizeControl();
        }

        private void OnResize(object sender, EventArgs e) { ResizeControl(); }

        private void ResizeControl()
        {
            int width = ClientRectangle.Width - 2 * b;

            int y = b;
            Composant.Location = new System.Drawing.Point(b, y);
            Composant.Size = new Size(width, Composant.Height);

            y += j + Composant.Height + 8;
            int x = b;
            Groupe1.Location = new System.Drawing.Point(x, y);

            x += j + Groupe1.Width;
            Groupe2.Location = new System.Drawing.Point(x, y);
            Groupe2.Size = new Size(Groupe2.Width, Groupe2.Height);

            y += j + Math.Max(Groupe1.Height, Groupe2.Height);
            Scanner.Location = new System.Drawing.Point(b, y);
            Scanner.Size = new Size(width, Scanner.Height);

            y += j + Scanner.Height;
            Importer.Location = new System.Drawing.Point(b, y);
            Importer.Size = new Size(width, Importer.Height);

            y += j + Importer.Height;
            MaJ.Location = new System.Drawing.Point(b, y);
            MaJ.Size = new Size(width, MaJ.Height);

            y += j + MaJ.Height;
            EditParametre.Location = new System.Drawing.Point(b, y);
            EditParametre.Size = new Size(width - j - ValidParam.Width, EditParametre.Height);
            ValidParam.Location = new System.Drawing.Point(b + EditParametre.Width + j, y - 2);

            y += j + EditParametre.Height;
            ListBoxParams.Location = new System.Drawing.Point(b, y);
            ListBoxParams.Size = new Size(width, ListBoxParams.Height);

            y += 3 * j + ListBoxParams.Height + 8;
            EditPropriete.Location = new System.Drawing.Point(b, y);
            EditPropriete.Size = new Size(width - j - ValidProp.Width, EditPropriete.Height);
            ValidProp.Location = new System.Drawing.Point(b + EditPropriete.Width + j, y - 2);

            y += j + EditPropriete.Height;
            int h = ClientRectangle.Height - y - j - b;
            ListBoxPropMdl.Location = new System.Drawing.Point(b, y);
            ListBoxPropMdl.Size = new Size(width, (int)Math.Floor(h * 0.6));

            y += j + ListBoxPropMdl.Height;
            ListBoxPropCfg.Location = new System.Drawing.Point(b, y);
            ListBoxPropCfg.Size = new Size(width, Math.Max(10, ClientRectangle.Height - y - b));
        }

        private const String NomFichierParam = "Parametres.txt";
        private DicParam _DicParam = null;
        private ModelDoc2 MdlActif = null;
        private String Chemin = "";

        public int ActiveDocChange()
        {
            MdlActif = _Sw.ActiveDoc;

            if (MdlActif.IsRef())
                Rechercher_Fichier();
            else
                ViderParam();

            return 1;
        }

        #region Parametres

        private void ViderParam()
        {
            Chemin = "";
            Composant.Text = "Aucun";
            _DicParam.Vider();
            ViderModele();
        }

        private void AfficherComposant()
        {
            Composant.Text = String.Format("{0} [{1}]", MdlActif.eNomSansExt(), String.Join("/", MdlActif.eDossier().Split('\\').Reverse()));
        }

        private void Rechercher_Fichier()
        {
            if (!App.MacroEnCours)
            {
                try
                {
                    String OldChemin = Chemin;
                    // On regarde dans le repertoire du modèle courant

                    Chemin = Path.Combine(MdlActif.eDossier(), NomFichierParam);

                    AfficherComposant();

                    if (Chemin == OldChemin)
                        return;

                    ViderParam();

                    AfficherComposant();

                    if (!File.Exists(Chemin))
                    {
                        Chemin = Path.Combine(Directory.GetParent(MdlActif.eDossier()).FullName, NomFichierParam);
                        if (!File.Exists(Chemin))
                        {
                            Chemin = Path.Combine(MdlActif.eDossier(), NomFichierParam);
                        }
                    }

                    if (!(String.IsNullOrWhiteSpace(Chemin) || !File.Exists(Chemin)))
                        using (StreamReader sr = new StreamReader(Chemin, Texte.eGetEncoding(Chemin)))
                        {
                            while (sr.Peek() > -1)
                                _DicParam.Ajouter(sr.ReadLine());
                        }

                }
                catch (Exception e)
                {
                    //this.LogMethode(new Object[] { e });
                    ViderParam();
                }
            }
            else
            {
                ViderParam();
            }
        }

        private Boolean FichierCharger
        {
            get
            {
                ModelDoc2 m = _Sw.ActiveDoc;

                if (m.IsRef())
                {
                    if (String.IsNullOrWhiteSpace(Chemin))
                    {
                        Rechercher_Fichier();

                        if (!String.IsNullOrWhiteSpace(Chemin))
                            return true;
                    }
                    else
                        return true;
                }

                ViderParam();
                return false;

            }
        }

        private Boolean modifParam = false;

        private void ValidParam_Click(object sender, EventArgs e)
        {
            if (!FichierCharger) return;

            modifParam = true;

            try
            {
                String ligne = EditParametre.Text;
                String nvxParamNom = "";
                String nvxParamVal = "";

                String[] nvxtab = ligne.Split(new Char[] { ':' }, StringSplitOptions.None);
                if (nvxtab.Length == 2)
                {
                    nvxParamNom = nvxtab[0].Trim();
                    nvxParamVal = nvxtab[1].Trim();

                    if (_DicParam.Contains(nvxParamNom))
                    {
                        _DicParam[nvxParamNom].Expression = nvxParamVal;
                        _DicParam.Calculer();
                    }
                    else
                        _DicParam.Ajouter(ligne);

                    SauverFichier();
                }
            }
            catch (Exception ex)
            { this.LogMethode(new Object[] { ex }); }

            modifParam = false;
        }

        private void SauverFichier()
        {
            if (!FichierCharger) return;

            if (!File.Exists(Chemin))
            {
                StreamWriter sw = File.CreateText(Chemin);
                sw.Close();
            }

            using (StreamWriter sw = new StreamWriter(Chemin, false, Texte.eGetEncoding(Chemin)))
            {
                sw.WriteAsync(_DicParam.Texte());
            }
        }

        private enum action_e
        {
            Importer,
            MettreAjour
        }

        private void Importer_Click(object sender, EventArgs e)
        {
            List<Param> ListeParam = new List<Param>();

            foreach (int index in ListBoxParams.SelectedIndices)
                ListeParam.Add(_DicParam[index]);

            Appliquer(ListeParam, action_e.Importer);

            ViderModele();
            Rechercher_Propriete_Modele();
        }

        private void Maj_Click(object sender, EventArgs e)
        {
            List<Param> ListeParam = new List<Param>();

            foreach (Param item in _DicParam)
                ListeParam.Add(item);

            Appliquer(ListeParam, action_e.MettreAjour);

            ViderModele();
            Rechercher_Propriete_Modele();
        }

        private void Appliquer(List<Param> ListeParam, action_e action)
        {
            if (!FichierCharger) return;

            Dictionary<String, Component2> DicCp = new Dictionary<String, Component2>();

            Func<Component2, String> Cle = delegate (Component2 cp)
            {
                String cle = cp.GetPathName();
                if (radioMdl.Checked)
                    return cle;

                return cle + "_" + cp.eNomConfiguration();
            };

            Action<Component2, Param> AppliquerProp = delegate (Component2 cp, Param p)
            {
                ModelDoc2 mdl = cp.eModelDoc2();
                String cle = p.Nom, val = p.Resultat;

                String cfg = "";
                if (radioCfg.Checked)
                    cfg = cp.eNomConfiguration();

                if (mdl.ePropExiste(cle, cfg))
                {
                    mdl.eGestProp(cfg).ePropSet(cle, val);
                    return;
                }

                if (action == action_e.Importer)
                    mdl.eGestProp(cfg).ePropAdd(cle, val);
            };

            Component2 CpCourant = MdlActif.eComposantRacine(false);
            Component2 CpEdited = null;

            if (MdlActif.TypeDoc() == eTypeDoc.Assemblage)
                CpEdited = MdlActif.eAssemblyDoc().GetEditTargetComponent();

            if(CpEdited.IsRef() && (CpCourant.GetPathName() != CpEdited.GetPathName()))
                DicCp.Add(Cle(CpEdited) , CpEdited);
            else
                DicCp.Add("Courant", CpCourant);

            // Si c'est un assemblage, on liste les composants
            if (radioTtModeles.Checked && (MdlActif.TypeDoc() == eTypeDoc.Assemblage))
                MdlActif.eRecParcourirComposants(
                c =>
                {
                    String cle = Cle(c);
                    if (!DicCp.ContainsKey(cle) && c.eEstDansLeDossier(MdlActif) && !c.IsHidden(true))
                        DicCp.Add(cle, c);

                    return false;
                },
                null
                );

            try
            {
                foreach (Component2 cp in DicCp.Values)
                {
                    foreach (Param k in ListeParam)
                        AppliquerProp(cp, k);
                }
            }
            catch (Exception ex)
            { this.LogMethode(new Object[] { ex }); }
        }

        private void ListeBoxParams_SelectionChanged(object sender, EventArgs e)
        {
            if (modifParam || !FichierCharger) return;

            if (ListBoxParams.SelectedIndex == ListBox.NoMatches) return;

            EditParametre.Text = _DicParam[ListBoxParams.SelectedIndex].SauvegardeFichier;
        }

        private void ListBoxParams_MouseDown(object sender, MouseEventArgs e)
        {
            int index = ListBoxParams.IndexFromPoint(e.Location);
            if (index != ListBox.NoMatches)
            {
                ListBoxParams.SelectedIndex = index;
                ListeBoxParams_SelectionChanged(null, null);
            }

            if (e.Button == MouseButtons.Right)
            {
                if (index != ListBox.NoMatches)
                {
                    ParamContextMenuListBox.Show(Cursor.Position);
                    ParamContextMenuListBox.Visible = true;
                }
                else
                    ParamContextMenuListBox.Visible = false;
            }
            else if ((e.Button == MouseButtons.Left) && (index != ListBox.NoMatches))
            {
                ListBoxParams.DoDragDrop(ListBoxParams.SelectedItem, DragDropEffects.Move);
            }
        }

        private void ListBoxParams_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void ListBoxParams_DragDrop(object sender, DragEventArgs e)
        {
            System.Drawing.Point point = ListBoxParams.PointToClient(new System.Drawing.Point(e.X, e.Y));
            int To = ListBoxParams.IndexFromPoint(point);
            if (To < 0)
                To = ListBoxParams.Items.Count - 1;

            String data = (String)e.Data.GetData(typeof(String));
            int From = ListBoxParams.Items.IndexOf(data);

            if (To == From) return;

            modifParam = true;
            _DicParam.Move(From, To);
            modifParam = false;

            SauverFichier();
        }

        private void SupprimerParam(object sender, EventArgs e)
        {
            if (!FichierCharger) return;

            if (ListBoxParams.SelectedIndex == ListBox.NoMatches) return;

            modifParam = true;
            _DicParam.Effacer(ListBoxParams.SelectedIndex);
            EditParametre.Text = "";
            modifParam = false;

            SauverFichier();
        }

        private void Scanner_Click(object sender, EventArgs e)
        {
            if (!FichierCharger) return;

            Dictionary<String, Component2> DicCp = new Dictionary<String, Component2>();

            DicCp.Add("Courant", MdlActif.eComposantRacine(false));

            Func<Component2, String> Cle = delegate (Component2 cp)
            {
                String cle = cp.GetPathName();
                if (radioMdl.Checked)
                    return cle;

                return cle + "_" + cp.eNomConfiguration();
            };

            Action<Component2> ScannerProp = delegate (Component2 cp)
            {
                ModelDoc2 mdl = cp.eModelDoc2();

                String cfg = "";
                if (radioCfg.Checked)
                    cfg = cp.eNomConfiguration();

                var DicProp = mdl.eGestProp(cfg).eListProp();

                foreach (var nom in DicProp.Keys)
                {
                    if (!HashPropExclu.Contains(nom))
                        _DicParam.Ajouter(nom, DicProp[nom]);
                }

            };

            // Si c'est un assemblage, on liste les composants
            if (radioTtModeles.Checked && (MdlActif.TypeDoc() == eTypeDoc.Assemblage))
                MdlActif.eRecParcourirComposants(
                c =>
                {
                    String cle = Cle(c);
                    if (!DicCp.ContainsKey(cle) && c.eEstDansLeDossier(MdlActif) && !c.IsHidden(true))
                        DicCp.Add(cle, c);

                    return false;
                },
                null
                );

            try
            {
                foreach (Component2 cp in DicCp.Values)
                    ScannerProp(cp);

                SauverFichier();
            }
            catch (Exception ex)
            { this.LogMethode(new Object[] { ex }); }


        }

        private class Param
        {
            private String _Nom = "";
            private String _Expression = "";
            private Double? _Valeur = null;

            public String Nom { get { return _Nom; } private set { _Nom = value; } }

            public String Expression
            {
                get
                {
                    return _Expression;
                }

                set
                {
                    int _Index;
                    try
                    { _Index = _Dic.ListBox.Items.IndexOf(AfficheListBox); }
                    catch
                    { _Index = -1; }

                    _Expression = value;
                    _Valeur = _Expression.Evaluer();

                    if (Valeur.IsRef())
                    {
                        if (Calcul.Variables.ContainsKey(Nom))
                            Calcul.Variables[Nom] = (Double)Valeur;
                        else
                            Calcul.Variables.Add(Nom, (Double)Valeur);
                    }

                    if (_Index > -1)
                        _Dic.ListBox.Items[_Index] = AfficheListBox;
                }
            }

            public void Calculer()
            {
                Expression = _Expression;
            }

            public Double? Valeur
            {
                get { return _Valeur; }
            }

            public String Resultat
            {
                get
                {
                    String r = Expression;

                    if (Valeur.IsRef())
                        r = _Valeur.ToString();

                    return r;
                }
            }

            private DicParam _Dic = null;

            public Boolean Valid { get; private set; }

            public Param(DicParam dic, String ligne)
            {
                _Dic = dic;
                String[] tab = ligne.Split(new Char[] { ':' }, StringSplitOptions.None);
                if (tab.Length == 2)
                {
                    Nom = tab[0].Trim();

                    Expression = tab[1].Trim();
                    AjouterDic();
                }
            }

            public Param(DicParam dic, String nom, String expression)
            {
                _Dic = dic;
                Nom = nom;

                Expression = expression;
                AjouterDic();
            }

            private void AjouterDic()
            {
                Valid = false;

                if (!_Dic.Contains(Nom))
                {
                    _Dic.Add(this);
                    _Dic.ListBox.Items.Add(AfficheListBox);

                    Valid = true;
                }
            }

            public String SauvegardeFichier
            {
                get
                {
                    return _Nom + " : " + _Expression;
                }
            }

            public String AfficheListBox
            {
                get
                {
                    if (Valeur.IsRef())
                        return SauvegardeFichier + " = " + Math.Round((Double)Valeur, 3).ToString();

                    return SauvegardeFichier;
                }
            }
        }

        private class DicParam : KeyedCollection<String, Param>
        {
            public ListBox ListBox;
            public DicParam(ListBox listBox)
            {
                ListBox = listBox;
            }

            public Boolean Ajouter(String ligne)
            {
                Param param = new Param(this, ligne);
                return ValiderParam(param);
            }

            public Boolean Ajouter(String nom, String expression)
            {
                Param param = new Param(this, nom, expression);
                return ValiderParam(param);
            }

            private Boolean ValiderParam(Param p)
            {
                if (p.Valid)
                    return true;

                p = null;
                return false;
            }

            public void Vider()
            {
                Clear();
                ListBox.Items.Clear();
                Calcul.Variables.Reinitialiser();
            }

            protected override string GetKeyForItem(Param item)
            {
                return item.Nom;
            }

            public String Texte()
            {
                List<String> l = new List<string>();
                foreach (var item in this)
                {
                    l.Add(item.SauvegardeFichier);
                }

                return String.Join("\r\n", l);
            }

            public void Move(int From, int To)
            {
                String data = (String)ListBox.Items[From];
                Param param = this[From];
                this.Remove(param);
                this.Insert(To, param);

                ListBox.Items.Remove(data);
                ListBox.Items.Insert(To, data);

                Calcul.Variables.Reinitialiser();

                Calculer();
            }

            public void Effacer(int i)
            {
                RemoveAt(i);
                ListBox.Items.RemoveAt(i);

                Calcul.Variables.Reinitialiser();

                Calculer();
            }

            public void Calculer()
            {
                foreach (Param p in this)
                    p.Calculer();
            }
        }

        #endregion

        #region Propriete

        private void ViderModele()
        {
            EditPropriete.Text = "";
            _DicPropModele.Vider();
            _DicPropCfg.Vider();
            EditedProp = null;
        }

        private DicProp _DicPropModele = null;
        private DicProp _DicPropCfg = null;

        public int Rechercher_Propriete_Modele()
        {
            ModelDoc2 mdl = _Sw.ActiveDoc;

            if (mdl.IsNull()) return 1;

            Dictionary<String, String> dic = mdl.eListProp();

            foreach (String nom in dic.Keys)
            {
                if (!HashPropExclu.Contains(nom))
                    _DicPropModele.Ajouter(nom, dic[nom], mdl);
            }

            switch (mdl.TypeDoc())
            {
                case eTypeDoc.Piece:
                    PartDoc p = mdl.ePartDoc();
                    p.ActiveConfigChangePostNotify -= Rechercher_Propriete_Config;
                    p.ActiveConfigChangePostNotify += Rechercher_Propriete_Config;
                    break;
                case eTypeDoc.Assemblage:
                    AssemblyDoc a = mdl.eAssemblyDoc();
                    a.ActiveConfigChangePostNotify -= Rechercher_Propriete_Config;
                    a.ActiveConfigChangePostNotify += Rechercher_Propriete_Config;
                    break;
                default:
                    break;
            }

            Rechercher_Propriete_Config();

            return 1;
        }

        public int Rechercher_Propriete_Config()
        {
            ModelDoc2 mdl = _Sw.ActiveDoc;

            _DicPropCfg.Vider();

            if (mdl.IsNull()) return 1;

            String cfg = mdl.eNomConfigActive();
            Dictionary<String, String> dic = mdl.eListProp(cfg);

            foreach (String nom in dic.Keys)
            {
                if (!HashPropExclu.Contains(nom))
                    _DicPropCfg.Ajouter(nom, dic[nom], mdl, cfg);
            }

            return 1;
        }

        private Boolean modifProp = false;

        private void EffacerEditPropriete()
        {
            EditedProp = null;
            EditPropriete.Text = "";
        }

        private Prop EditedProp = null;

        private void ValiProp_Click(object sender, EventArgs e)
        {
            if (!FichierCharger) return;

            modifProp = true;

            try
            {
                String ligne = EditPropriete.Text;
                String nvxParamNom = "";
                String nvxParamVal = "";

                String[] nvxtab = ligne.Split(new Char[] { ':' }, StringSplitOptions.None);
                if (nvxtab.Length == 2)
                {
                    nvxParamNom = nvxtab[0].Trim();
                    nvxParamVal = nvxtab[1].Trim();

                    if (!String.IsNullOrWhiteSpace(nvxParamVal) && EditedProp.IsRef())
                    {
                        EditedProp.Valeur = nvxParamVal;
                        EditedProp.Sauver();
                    }
                }
            }
            catch (Exception ex)
            { this.LogMethode(new Object[] { ex }); }

            modifProp = false;
        }

        private void ListBoxProp_SelectionChanged(object sender, EventArgs e)
        {
            ListBox box = (ListBox)sender;

            if (modifProp || !FichierCharger) return;

            if (box.SelectedIndex == ListBox.NoMatches) return;

            EditedProp = DicPropFromBox(box)[box.SelectedIndex];
            EditPropriete.Text = EditedProp.AfficheListBox;
        }

        private void ListBoxProp_MouseDown(object sender, MouseEventArgs e)
        {
            ListBox box = (ListBox)sender;

            int index = box.IndexFromPoint(e.Location);
            if (index != ListBox.NoMatches)
            {
                box.SelectedIndex = index;
                ListBoxProp_SelectionChanged(box, null);
            }

            if (e.Button == MouseButtons.Right)
            {
                if (index != ListBox.NoMatches)
                {
                    ParamContextMenuListBox.Show(Cursor.Position);
                    ParamContextMenuListBox.Visible = true;
                }
                else
                    ParamContextMenuListBox.Visible = false;
            }
        }

        private void SupprimerProp(object sender, EventArgs e)
        {
            ListBox box = (ListBox)((ContextMenuStrip)(((ToolStripMenuItem)sender).Owner)).SourceControl;

            if (box.SelectedIndex == ListBox.NoMatches) return;

            DicPropFromBox(box).EffacerProp(box.SelectedIndex);
            EffacerEditPropriete();
        }

        private DicProp DicPropFromBox(ListBox box)
        {
            DicProp dic = _DicPropCfg;
            if (box == _DicPropModele.ListBox)
                dic = _DicPropModele;

            return dic;
        }

        private class Prop
        {
            private String _Nom = "";
            private String _Valeur = "";

            private ModelDoc2 mdl = null;
            private String config = "";

            public ModelDoc2 Modele
            {
                get { return mdl; }
            }

            public String Config
            {
                get { return config; }
            }

            public String Nom { get { return _Nom; } private set { _Nom = value; } }

            public String Valeur
            {
                get
                {
                    return _Valeur;
                }

                set
                {
                    int _Index;
                    try
                    { _Index = _Dic.ListBox.Items.IndexOf(AfficheListBox); }
                    catch
                    { _Index = -1; }

                    _Valeur = value;

                    if (_Index > -1)
                        _Dic.ListBox.Items[_Index] = AfficheListBox;
                }
            }

            private DicProp _Dic = null;

            public Boolean Valid { get; private set; }

            public Prop(DicProp Dic, String Nom, String Val, ModelDoc2 Modele, String Config = "")
            {
                _Dic = Dic;

                this.Nom = Nom;
                mdl = Modele;
                config = Config;

                Valeur = Val;

                AjouterDic();
            }

            private void AjouterDic()
            {
                Valid = false;

                if (!_Dic.Contains(Nom))
                {
                    _Dic.Add(this);
                    _Dic.ListBox.Items.Add(AfficheListBox);

                    Valid = true;
                }
            }

            public String AfficheListBox
            {
                get
                {
                    return _Nom + " : " + _Valeur;
                }
            }

            public void Sauver()
            {
                mdl.eGestProp(config).ePropSet(Nom, Valeur);
            }
        }

        private class DicProp : KeyedCollection<String, Prop>
        {
            public ListBox ListBox;
            public DicProp(ListBox listBox)
            {
                ListBox = listBox;
            }

            public Boolean Ajouter(String Nom, String Val, ModelDoc2 Mdl, String Config = "")
            {
                Prop prop = new Prop(this, Nom, Val, Mdl, Config);
                if (prop.Valid)
                    return true;

                prop = null;
                return false;
            }

            public void Vider()
            {
                Clear();
                ListBox.Items.Clear();
            }

            protected override string GetKeyForItem(Prop item)
            {
                return item.Nom;
            }

            public void EffacerProp(int i)
            {
                Prop p = this[i];
                p.Modele.eGestProp(p.Config).Delete2(p.Nom);
                RemoveAt(i);
                ListBox.Items.RemoveAt(i);
            }
        }

        #endregion
    }
}
