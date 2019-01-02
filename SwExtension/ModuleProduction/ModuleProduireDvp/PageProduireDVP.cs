using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ModuleProduction.ModuleProduireDvp
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Creer les dvp de tôle"),
        ModuleNom("ProduireDvp"),
        ModuleDescription("Creer les développés des tôles.")
        ]
    public class PageProduireDvp : BoutonPMPManager
    {
        private Parametre ConvertirEsquisse;
        private Parametre QuantiteDiff;
        private Parametre PropQuantite;
        private Parametre AfficherLignePliage;
        private Parametre AfficherNotePliage;

        private Parametre InscrireNomTole;
        private Parametre TailleInscription;
        private Parametre FermerPlan;
        private Parametre OrienterDvp;
        private Parametre OrientationDvp;


        public PageProduireDvp()
        {
            try
            {
                PropQuantite = _Config.AjouterParam("PropQuantite", CONSTANTES.PROPRIETE_QUANTITE, "Propriete \"Quantite\"", "Recherche cette propriete");
                QuantiteDiff = _Config.AjouterParam("QuantiteDiff", true, "Calculer la différence");
                AfficherLignePliage = _Config.AjouterParam("AfficherLignePliage", true, "Afficher les lignes de pliage");
                AfficherNotePliage = _Config.AjouterParam("AfficherNotePliage", true, "Afficher les notes de pliage");
                InscrireNomTole = _Config.AjouterParam("InscrireNomTole", true, "Inscrire la réf du dvp sur la tole");
                TailleInscription = _Config.AjouterParam("TailleInscription", 5, "Ht des inscriptions en mm", "Ht des inscriptions en mm");

                OrienterDvp = _Config.AjouterParam("OrienterDvp", false, "Orienter les dvps");
                OrientationDvp = _Config.AjouterParam("OrientationDvp", eOrientation.Portrait);
                FermerPlan = _Config.AjouterParam("FermerPlan", false, "Fermer les plans", "Fermer les plans après génération des dvps");
                ConvertirEsquisse = _Config.AjouterParam("ConvertirEsquisse", false, "Convertir les dvp en esquisse", "Le lien entre le dvp et le modèle est cassé, la config dvp est supprimée après la création de la vue");

                OnCalque += Calque;
                OnRunAfterActivation += Rechercher_Infos;
                OnRunOkCommand += RunOkCommand;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private ModelDoc2 MdlBase;

        private CtrlTextBox _Texte_RefFichier;
        private CtrlTextBox _TextBox_Campagne;
        private CtrlCheckBox _CheckBox_Quantite_Diff;
        private CtrlTextBox _Texte_Quantite;
        private CtrlTextListBox _TextListBox_Materiaux;
        private CtrlTextListBox _TextListBox_Ep;
        private CtrlTextBox _Texte_TailleInscription;
        private CtrlCheckBox _CheckBox_MettreAjourCampagne;
        private CtrlCheckBox _CheckBox_AfficherLignePliage;
        private CtrlCheckBox _CheckBox_AfficherNotePliage;
        private CtrlCheckBox _CheckBox_InscrireNomTole;
        private CtrlCheckBox _CheckBox_OrienterDvp;
        private CtrlEnumComboBox<eOrientation, Intitule> _EnumComboBox_OrientationDvp;

        protected void Calque()
        {
            try
            {
                MdlBase = App.ModelDoc2;

                Groupe G;

                G = _Calque.AjouterGroupe("Fichier");

                _Texte_RefFichier = G.AjouterTexteBox("Référence du fichier :", "la référence est ajoutée au début du nom de chaque fichier généré");

                _Texte_RefFichier.Text = Ref;
                _Texte_RefFichier.LectureSeule = false;

                // S'il n'y a pas de reference, on met le texte en rouge
                if (String.IsNullOrWhiteSpace(Ref))
                    _Texte_RefFichier.BackgroundColor(Color.Red, true);

                _TextBox_Campagne = G.AjouterTexteBox("Campagne :", "");
                _TextBox_Campagne.LectureSeule = true;

                _CheckBox_MettreAjourCampagne = G.AjouterCheckBox("Mettre à jour la campagne");

                G = _Calque.AjouterGroupe("Quantité :");

                _CheckBox_Quantite_Diff = G.AjouterCheckBox(QuantiteDiff);

                _Texte_Quantite = G.AjouterTexteBox("Multiplier par quantité :", "Multiplier les quantités par");
                _Texte_Quantite.Text = Quantite();
                _Texte_Quantite.ValiderTexte += ValiderTextIsInteger;

                G = _Calque.AjouterGroupe("Materiaux :");

                _TextListBox_Materiaux = G.AjouterTextListBox();
                _TextListBox_Materiaux.TouteHauteur = true;
                _TextListBox_Materiaux.Height = 50;
                _TextListBox_Materiaux.SelectionMultiple = true;

                G = _Calque.AjouterGroupe("Ep :");

                _TextListBox_Ep = G.AjouterTextListBox();
                _TextListBox_Ep.TouteHauteur = true;
                _TextListBox_Ep.Height = 50;
                _TextListBox_Ep.SelectionMultiple = true;

                G = _Calque.AjouterGroupe("Options");

                _CheckBox_AfficherLignePliage = G.AjouterCheckBox(AfficherLignePliage);
                _CheckBox_AfficherNotePliage = G.AjouterCheckBox(AfficherNotePliage);
                _CheckBox_AfficherNotePliage.StdIndent();

                _CheckBox_AfficherLignePliage.OnUnCheck += _CheckBox_AfficherNotePliage.UnCheck;
                _CheckBox_AfficherLignePliage.OnIsCheck += _CheckBox_AfficherNotePliage.IsEnable;

                // Pour eviter d'ecraser le parametre de "AfficherNotePliage", le met à jour seulement si
                if (!_CheckBox_AfficherLignePliage.IsChecked)
                {
                    _CheckBox_AfficherNotePliage.IsChecked = _CheckBox_AfficherLignePliage.IsChecked;
                    _CheckBox_AfficherNotePliage.IsEnabled = _CheckBox_AfficherLignePliage.IsChecked;
                }

                _CheckBox_InscrireNomTole = G.AjouterCheckBox(InscrireNomTole);
                _Texte_TailleInscription = G.AjouterTexteBox(TailleInscription, false);
                _Texte_TailleInscription.ValiderTexte += ValiderTextIsInteger;
                _Texte_TailleInscription.StdIndent();

                _CheckBox_InscrireNomTole.OnIsCheck += _Texte_TailleInscription.IsEnable;
                _Texte_TailleInscription.IsEnabled = _CheckBox_InscrireNomTole.IsChecked;


                _CheckBox_OrienterDvp = G.AjouterCheckBox(OrienterDvp);
                _EnumComboBox_OrientationDvp = G.AjouterEnumComboBox<eOrientation, Intitule>(OrientationDvp);
                _EnumComboBox_OrientationDvp.StdIndent();

                _CheckBox_OrienterDvp.OnIsCheck += _EnumComboBox_OrientationDvp.IsEnable;
                _EnumComboBox_OrientationDvp.IsEnabled = _CheckBox_OrienterDvp.IsChecked;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }
        private String _Ref = "";

        private String Ref
        {
            get
            {
                if (String.IsNullOrWhiteSpace(_Ref))
                    _Ref = MdlBase.eRefFichierComplet();

                return _Ref;
            }
        }

        private String Quantite()
        {
            CustomPropertyManager PM = MdlBase.Extension.get_CustomPropertyManager("");

            if (MdlBase.ePropExiste(PropQuantite.GetValeur<String>()))
            {
                return Math.Max(MdlBase.eProp(PropQuantite.GetValeur<String>()).eToInteger(), 1).ToString();
            }

            return "1";
        }

        private int Campagne;
        private List<String> ListeMateriaux;
        private List<String> ListeEp;
        private SortedDictionary<int, Corps> ListeCorps;

        protected void Rechercher_Infos()
        {
            try
            {
                WindowLog.Ecrire("Recherche des materiaux et epaisseurs ");

                ListeCorps = MdlBase.pChargerNomenclature(eTypeCorps.Tole);
                ListeMateriaux = new List<String>();
                ListeEp = new List<String>();
                Campagne = 1;

                foreach (var corps in ListeCorps.Values)
                {
                    if (corps.TypeCorps != eTypeCorps.Tole) continue;

                    Campagne = Math.Max(Campagne, corps.Campagne.Keys.Max());

                    ListeMateriaux.AddIfNotExist(corps.Materiau);
                    ListeEp.AddIfNotExist(corps.Dimension);
                }

                WindowLog.SautDeLigne();

                ListeMateriaux.Sort(new WindowsStringComparer());
                ListeEp.Sort(new WindowsStringComparer());

                _TextBox_Campagne.Text = Campagne.ToString();
                _TextListBox_Materiaux.Liste = ListeMateriaux;
                _TextListBox_Materiaux.ToutSelectionner(false);

                _TextListBox_Ep.Liste = ListeEp;
                _TextListBox_Ep.ToutSelectionner(false);

                if (Campagne == 1)
                {
                    _CheckBox_Quantite_Diff.IsEnabled = false;
                    _CheckBox_Quantite_Diff.Visible = false;
                }

                if (!File.Exists(Path.Combine(MdlBase.pDossierLaserTole(), Campagne.ToString(), CONST_PRODUCTION.FICHIER_NOMENC)))
                {
                    _CheckBox_MettreAjourCampagne.IsEnabled = false;
                    _CheckBox_MettreAjourCampagne.Visible = false;
                }
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdProduireDvp Cmd = new CmdProduireDvp();

            Cmd.MdlBase = App.Sw.ActiveDoc;
            Cmd.ListeCorps = ListeCorps;
            Cmd.RefFichier = _Texte_RefFichier.Text.Trim();
            Cmd.Quantite_Diff = _CheckBox_Quantite_Diff.IsChecked;
            Cmd.Quantite = _Texte_Quantite.Text.eToInteger();
            Cmd.IndiceCampagne = _TextBox_Campagne.Text.eToInteger();
            Cmd.MettreAjourCampagne = _CheckBox_MettreAjourCampagne.IsChecked;
            Cmd.ListeMateriaux = _TextListBox_Materiaux.ListSelectedText.Count > 0 ? _TextListBox_Materiaux.ListSelectedText : _TextListBox_Materiaux.Liste;
            Cmd.ListeEp = _TextListBox_Ep.ListSelectedText.Count > 0 ? _TextListBox_Ep.ListSelectedText : _TextListBox_Ep.Liste;
            Cmd.AfficherLignePliage = _CheckBox_AfficherLignePliage.IsChecked;
            Cmd.AfficherNotePliage = _CheckBox_AfficherNotePliage.IsChecked;
            Cmd.InscrireNomTole = _CheckBox_InscrireNomTole.IsChecked;
            Cmd.TailleInscription = _Texte_TailleInscription.Text.eToInteger();
            Cmd.OrienterDvp = _CheckBox_OrienterDvp.IsChecked;
            Cmd.OrientationDvp = _EnumComboBox_OrientationDvp.Val;

            Cmd.Executer();
        }
    }
}
