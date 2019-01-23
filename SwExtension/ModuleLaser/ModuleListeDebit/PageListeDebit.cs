using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Xml;

namespace ModuleLaser
{
    namespace ModuleListeDebit
    {
        [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
            ModuleTitre("Liste de débit"),
            ModuleNom("ListeDebit"),
            ModuleDescription("Création de la liste de débit.")
            PageOptions(swPropertyManagerPageOptions_e.swPropertyManagerOptions_MultiplePages)
            ]
        public class PageListeDebit : BoutonPMPManager
        {
            private Parametre PropQuantite;
            private Parametre PrendreEnCompteTole;

            private Parametre ComposantsExterne;

            private Parametre LgBarre;

            private Parametre AfficherListe;

            public PageListeDebit()
            {
                try
                {
                    PropQuantite = _Config.AjouterParam("PropQuantite", CONSTANTES.PROPRIETE_QUANTITE, "Propriete \"Quantite\"", "Recherche cette propriete");
                    PrendreEnCompteTole = _Config.AjouterParam("PrendreEnCompteTole", true, "Prendre en compte les tôles");
                    ComposantsExterne = _Config.AjouterParam("ComposantExterne", false, "Prendre en compte les composants externe au dossier du modèle");
                    LgBarre = _Config.AjouterParam("LgBarre", 6000, "Lg d'une barre");
                    AfficherListe = _Config.AjouterParam("AfficherListe", true, "Afficher la liste après execution");

                    OnCalque += Calque;
                    OnNextPage += RunOnNextPage;
                    OnPreviousPage += delegate { return false; };
                    OnRunAfterActivation += Rechercher_Materiaux;
                    OnRunOkCommand += RunOkCommand;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private List<Groupe> ListeGroupe1 = new List<Groupe>();
            private List<Groupe> ListeGroupe2 = new List<Groupe>();
            private List<CtrlCheckBox> ListeCheckBoxLg = new List<CtrlCheckBox>();
            private List<CtrlTextBox> ListeTextBoxLg = new List<CtrlTextBox>();

            private CtrlTextBox _Texte_RefFichier;
            private CtrlTextListBox _TextListBox_Materiaux;
            private CtrlTextBox _Texte_Quantite;
            private CtrlCheckBox _CheckBox_ComposantsExterne;
            private CtrlCheckBox _CheckBox_PrendreEnCompteTole;
            private CtrlCheckBox _CheckBox_ForcerMateriau;
            private CtrlTextComboBox _TextComboBox_ForcerMateriau;
            private CtrlOption _Option_ListeDebit;
            private CtrlOption _Option_ListeBarres;
            private CtrlTextBox _Texte_LgBarre;

            private CtrlCheckBox _CheckBox_AfficherListe;

            private eTypeSortie TypeSortie = eTypeSortie.ListeDebit;

            private readonly int NbProfilMax = 40;

            protected void Calque()
            {
                try
                {
                    Groupe G;

                    ListeGroupe1.Add(_Calque.AjouterGroupe("Fichier"));
                    G = ListeGroupe1.Last();

                    _Texte_RefFichier = G.AjouterTexteBox("Référence du fichier :", "la référence est ajoutée au début du nom de chaque fichier généré");

                    String Ref = MdlBase.eRefFichier();
                    _Texte_RefFichier.Text = Ref;
                    _Texte_RefFichier.LectureSeule = true;

                    // S'il n'y a pas de reference, on met le texte en rouge
                    if (String.IsNullOrWhiteSpace(Ref))
                        _Texte_RefFichier.BackgroundColor(Color.Red, true);

                    _Texte_Quantite = G.AjouterTexteBox("Quantité :", "Multiplier les quantités par");
                    _Texte_Quantite.Text = Quantite();
                    _Texte_Quantite.ValiderTexte += ValiderTextIsInteger;

                    _CheckBox_ComposantsExterne = G.AjouterCheckBox(ComposantsExterne);

                    ListeGroupe1.Add(_Calque.AjouterGroupe("Materiaux :"));
                    G = ListeGroupe1.Last();

                    _TextListBox_Materiaux = G.AjouterTextListBox();
                    _TextListBox_Materiaux.TouteHauteur = true;
                    _TextListBox_Materiaux.Height = 60;
                    _TextListBox_Materiaux.SelectionMultiple = true;

                    _CheckBox_ForcerMateriau = G.AjouterCheckBox("Forcer le materiau");
                    _TextComboBox_ForcerMateriau = G.AjouterTextComboBox();
                    _TextComboBox_ForcerMateriau.Editable = true;
                    _TextComboBox_ForcerMateriau.LectureSeule = false;
                    _TextComboBox_ForcerMateriau.NotifieSurSelection = false;
                    _TextComboBox_ForcerMateriau.IsEnabled = false;
                    _CheckBox_ForcerMateriau.OnIsCheck += _TextComboBox_ForcerMateriau.IsEnable;

                    ListeGroupe1.Add(_Calque.AjouterGroupe("Options"));
                    G = ListeGroupe1.Last();

                    _Option_ListeDebit = G.AjouterOption("Liste de débit");
                    _Option_ListeBarres = G.AjouterOption("Liste des barres");
                    _Option_ListeDebit.OnCheck += delegate (Object sender) { TypeSortie = eTypeSortie.ListeDebit; };
                    _Option_ListeBarres.OnCheck += delegate (Object sender) { TypeSortie = eTypeSortie.ListeBarre; };
                    _Option_ListeDebit.IsChecked = true;

                    _Texte_LgBarre = G.AjouterTexteBox(LgBarre, true);

                    _CheckBox_PrendreEnCompteTole = G.AjouterCheckBox(PrendreEnCompteTole);
                    _CheckBox_PrendreEnCompteTole.OnIsCheck += delegate (Object sender, Boolean value) { Rechercher_Materiaux(); };

                    _CheckBox_AfficherListe = G.AjouterCheckBox(AfficherListe);

                    ListeGroupe2.Add(_Calque.AjouterGroupe("Lg des barres"));
                    G = ListeGroupe2.Last();
                    G.Visible = false;

                    for (int i = 0; i < NbProfilMax; i++)
                    {
                        ListeCheckBoxLg.Add(G.AjouterCheckBox("Profil " + (i + 1)));
                        CtrlCheckBox c = ListeCheckBoxLg.Last();
                        c.Visible = false;

                        ListeTextBoxLg.Add(G.AjouterTexteBox());
                        CtrlTextBox t = ListeTextBoxLg.Last();
                        t.Text = _Texte_LgBarre.Text.eToInteger().ToString();
                        t.Visible = false;
                    }


                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private String Quantite()
            {
                CustomPropertyManager PM = MdlBase.Extension.get_CustomPropertyManager("");

                if (MdlBase.ePropExiste(PropQuantite.GetValeur<String>()))
                {
                    return Math.Max(MdlBase.eGetProp(PropQuantite.GetValeur<String>()).eToInteger(), 1).ToString();
                }

                return "1";
            }

            private List<String> ListeMateriaux;

            protected void Rechercher_Materiaux()
            {
                WindowLog.Ecrire("Recherche des materiaux : ");

                ListeMateriaux = MdlBase.ListeMateriaux(_CheckBox_PrendreEnCompteTole.IsChecked ? eTypeCorps.Tole | eTypeCorps.Barre : eTypeCorps.Barre);

                foreach (var m in ListeMateriaux)
                    WindowLog.Ecrire(" - " + m);

                WindowLog.SautDeLigne();

                _TextListBox_Materiaux.Liste = ListeMateriaux;
                _TextListBox_Materiaux.ToutSelectionner(false);
                _TextComboBox_ForcerMateriau.Liste = ListeMateriaux;
                _TextComboBox_ForcerMateriau.SelectedIndex = 0;
            }

            private CmdListeDebit Cmd = null;

            private CmdListeDebit.ListeLgProfil ListeLgProfil;

            private readonly String _TagRacine = "Base";
            private readonly String _TagMateriau = "Materiau";
            private readonly String _TagProfil = "Profil";
            private readonly String _TagNom = "name";

            private void SauverLg()
            {
                try
                {
                    // Création du Xml
                    XmlDocument xmlDoc = new XmlDocument();
                    //// On rajoute la déclaration
                    //XmlDeclaration xmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
                    //XmlElement root = xmlDoc.DocumentElement;
                    //xmlDoc.InsertBefore(xmlDeclaration, root);

                    // On ajoute le noeud de base
                    XmlNode Base = xmlDoc.CreateNode(XmlNodeType.Element, _TagRacine, "");
                    xmlDoc.AppendChild(Base);

                    foreach (var materiau in ListeLgProfil.DicLg.Keys)
                    {
                        XmlNode NdMateriau = xmlDoc.CreateNode(XmlNodeType.Element, _TagMateriau, "");
                        XmlAttribute AttMatt = xmlDoc.CreateAttribute(_TagNom);
                        AttMatt.Value = materiau;
                        NdMateriau.Attributes.SetNamedItem(AttMatt);
                        Base.AppendChild(NdMateriau);

                        foreach (var profil in ListeLgProfil.DicLg[materiau].Keys)
                        {
                            String lg = ListeLgProfil.DicLg[materiau][profil].ToString();

                            XmlNode NdProfil = xmlDoc.CreateNode(XmlNodeType.Element, _TagProfil, "");
                            XmlAttribute AttProfil = xmlDoc.CreateAttribute(_TagNom);
                            AttProfil.Value = profil;
                            NdProfil.Attributes.SetNamedItem(AttProfil);
                            NdProfil.InnerText = lg;

                            NdMateriau.AppendChild(NdProfil);
                        }
                    }

                    Log.Write(xmlDoc.OuterXml);

                    MdlBase.eSetListeLgProfils(xmlDoc.OuterXml);
                }
                catch
                { }
            }

            private void ChargerLg()
            {
                try
                {
                    String s = MdlBase.eGetListeLgProfils();
                    var DicMateriaux = ListeLgProfil.DicLg;

                    if (String.IsNullOrWhiteSpace(s) || DicMateriaux.IsNull()) return;

                    Func<XmlNode, String, String> GetTag = delegate (XmlNode n, String nomTag)
                    {
                        String Val = "";

                        if (n.Attributes[nomTag] != null)
                            Val = n.Attributes[nomTag].Value;

                        return Val;
                    };

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(s);

                    XmlNode Base = xmlDoc.SelectSingleNode(String.Format("/{0}", _TagRacine));

                    foreach (XmlNode nMateriau in Base.ChildNodes)
                    {
                        String materiau = GetTag(nMateriau, _TagNom);
                        if (DicMateriaux.ContainsKey(materiau))
                        {
                            var DicProfils = DicMateriaux[materiau];
                            foreach (XmlNode nProfil in nMateriau.ChildNodes)
                            {
                                String profil = GetTag(nProfil, _TagNom);
                                if (DicProfils.ContainsKey(profil))
                                {
                                    DicProfils[profil] = nProfil.InnerText.eToDouble();
                                }
                            }
                        }
                    }

                    xmlDoc = null;
                }
                catch
                { }
            }

            private Boolean ChargerCmd()
            {
                if (Cmd.IsRef()) return false;

                _Calque.CacherEntete();

                Cmd = new CmdListeDebit();

                Cmd.MdlBase = MdlBase;

                Cmd.TypeSortie = TypeSortie;
                Cmd.ListeMateriaux = _TextListBox_Materiaux.ListSelectedText.Count > 0 ? _TextListBox_Materiaux.ListSelectedText : _TextListBox_Materiaux.Liste;
                Cmd.ForcerMateriau = _CheckBox_ForcerMateriau.IsChecked ? _TextComboBox_ForcerMateriau.Text : null;
                Cmd.Quantite = _Texte_Quantite.Text.eToInteger();
                Cmd.PrendreEnCompteTole = _CheckBox_PrendreEnCompteTole.IsChecked;
                Cmd.ComposantsExterne = _CheckBox_ComposantsExterne.IsChecked;
                Cmd.RefFichier = _Texte_RefFichier.Text;
                Cmd.LgBarre = _Texte_LgBarre.Text.eToInteger();

                ListeLgProfil = Cmd.Analyser();

                ChargerLg();

                return true;
            }

            protected Boolean RunOnNextPage()
            {
                if (Cmd.IsRef()) return false;

                ChargerCmd();

                foreach (var groupe in ListeGroupe1)
                    groupe.Visible = false;

                foreach (var groupe in ListeGroupe2)
                    groupe.Visible = true;

                int i = 0;

                foreach (var materiau in ListeLgProfil.DicLg.Keys)
                {
                    foreach (var profil in ListeLgProfil.DicLg[materiau].Keys)
                    {
                        if (i == NbProfilMax) return true;

                        var c = ListeCheckBoxLg[i];
                        c.Visible = true;
                        c.Caption = String.Format("{0} [{1}]", profil, materiau);
                        c.IsChecked = true;

                        var t = ListeTextBoxLg[i++];
                        t.Visible = true;
                        t.Text = ListeLgProfil.DicLg[materiau][profil].ToString();
                        t.OnTextBoxChanged += delegate (Object sender, String text) { ListeLgProfil.DicLg[materiau][profil] = text.eToInteger(); };
                    }
                }

                return true;
            }

            protected void RunOkCommand()
            {
                if (!ChargerCmd())
                    SauverLg();

                WindowLog.Ecrire("Lg max des barres :");

                foreach (var materiau in ListeLgProfil.DicLg.Keys)
                {
                    foreach (var profil in ListeLgProfil.DicLg[materiau].Keys)
                    {
                        String lg = ListeLgProfil.DicLg[materiau][profil].ToString();
                        WindowLog.EcrireF("{0,10} [{1:10}] {2,10}", profil, materiau, lg);
                    }
                }

                Cmd.Executer();

                if(File.Exists(Cmd.CheminFichier))
                    System.Diagnostics.Process.Start(Cmd.CheminFichier);

            }
        }

        public static class SwAttributListeLgProfils
        {
            private const String ATTRIBUT_NOM = "ListeLgProfils";
            private const String ATTRIBUT_PARAM = "Val";

            private static AttributeDef AttDef = null;

            private static Parameter eAttributListeLgProfils(this ModelDoc2 mdl)
            {
                if (AttDef.IsNull())
                {
                    AttDef = App.Sw.DefineAttribute(ATTRIBUT_NOM);
                    AttDef.AddParameter(ATTRIBUT_PARAM, (int)swParamType_e.swParamTypeString, 0, 0);
                    AttDef.Register();
                }

                // Recherche de l'attribut dans la piece
                SolidWorks.Interop.sldworks.Attribute Att = null;
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

            public static void eSetListeLgProfils(this ModelDoc2 mdl, String val)
            {
                Parameter P = mdl.eAttributListeLgProfils();
                P.SetStringValue2(val, (int)swInConfigurationOpts_e.swAllConfiguration, "");
            }

            public static String eGetListeLgProfils(this ModelDoc2 mdl)
            {
                Parameter P = mdl.eAttributListeLgProfils();
                return P.GetStringValue();
            }
        }
    }
}
