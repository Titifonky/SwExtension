using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace ModuleLaser.ModuleCreerDvp
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Creer les dvp de tôle"),
        ModuleNom("CreerDvp"),
        ModuleDescription("Creer les développés des tôles.")
        ]
    public class PageCreerDvp : BoutonPMPManager
    {
        private Parametre ConvertirEsquisse;
        private Parametre PropQuantite;
        private Parametre AfficherLignePliage;
        private Parametre AfficherNotePliage;

        private Parametre ComposantsExterne;
        private Parametre InscrireNomTole;
        private Parametre TailleInscription;
        private Parametre FormatInscription;
        private List<String> ChampsInscription = new List<string>() { "<Nom_Piece>", "<Nom_Config>", "<No_Dossier>" };
        private Parametre FermerPlan;
        private Parametre OrienterDvp;
        private Parametre OrientationDvp;


        public PageCreerDvp()
        {
            try
            {
                PropQuantite = _Config.AjouterParam("PropQuantite", CONSTANTES.PROPRIETE_QUANTITE, "Propriete \"Quantite\"", "Recherche cette propriete");
                AfficherLignePliage = _Config.AjouterParam("AfficherLignePliage", true, "Afficher les lignes de pliage");
                AfficherNotePliage = _Config.AjouterParam("AfficherNotePliage", true, "Afficher les notes de pliage");
                InscrireNomTole = _Config.AjouterParam("InscrireNomTole", true, "Inscrire la réf du dvp sur la tole");
                TailleInscription = _Config.AjouterParam("TailleInscription", 5, "Ht des inscriptions en mm", "Ht des inscriptions en mm");
                FormatInscription = _Config.AjouterParam("FormatInscription", String.Join("-", ChampsInscription), "Format :", "Format de l'inscription à graver");

                ComposantsExterne = _Config.AjouterParam("ComposantExterne", false, "Créer les dvp des composants externe au dossier du modèle");
                OrienterDvp = _Config.AjouterParam("OrienterDvp", false, "Orienter les dvps");
                OrientationDvp = _Config.AjouterParam("OrientationDvp", eOrientation.Portrait);
                FermerPlan = _Config.AjouterParam("FermerPlan", false, "Fermer les plans", "Fermer les plans après génération des dvps");
                ConvertirEsquisse = _Config.AjouterParam("ConvertirEsquisse", false, "Convertir les dvp en esquisse", "Le lien entre le dvp et le modèle est cassé, la config dvp est supprimée après la création de la vue");

                OnCalque += Calque;
                OnRunAfterActivation += Rechercher_Materiaux;
                OnRunOkCommand += RunOkCommand;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private CtrlTextListBox _TextListBox_Materiaux;
        private CtrlCheckBox _CheckBox_ForcerMateriau;
        private CtrlTextComboBox _TextComboBox_ForcerMateriau;

        private CtrlTextBox _Texte_RefFichier;
        private CtrlTextBox _Texte_Quantite;
        private CtrlTextBox _Texte_TailleInscription;
        private CtrlTextBox _Texte_FormatInscription;
        private CtrlLabel _Label_InfoFormat;

        private CtrlCheckBox _CheckBox_ComposantsExterne;
        private CtrlCheckBox _CheckBox_AfficherLignePliage;
        private CtrlCheckBox _CheckBox_AfficherNotePliage;
        private CtrlCheckBox _CheckBox_InscrireNomTole;
        private CtrlCheckBox _CheckBox_OrienterDvp;
        private CtrlEnumComboBox<eOrientation, Intitule> _EnumComboBox_OrientationDvp;
        private CtrlCheckBox _CheckBox_FermerPlan;
        private CtrlCheckBox _CheckBox_ConvertirEsquisse;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Fichier");

                _Texte_RefFichier = G.AjouterTexteBox("Référence du fichier :", "la référence est ajoutée au début du nom de chaque fichier généré");

                String Ref = App.ModelDoc2.eRefFichier();
                _Texte_RefFichier.Text = Ref;
                _Texte_RefFichier.LectureSeule = false;

                // S'il n'y a pas de reference, on met le texte en rouge
                if (String.IsNullOrWhiteSpace(Ref))
                    _Texte_RefFichier.BackgroundColor(Color.Red, true);

                _Texte_Quantite = G.AjouterTexteBox("Quantité :", "Multiplier les quantités par");
                _Texte_Quantite.Text = Quantite();
                _Texte_Quantite.ValiderTexte += ValiderTextIsInteger;

                _CheckBox_ComposantsExterne = G.AjouterCheckBox(ComposantsExterne);

                G = _Calque.AjouterGroupe("Materiaux :");

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
                _Texte_FormatInscription = G.AjouterTexteBox(FormatInscription, true);
                _Texte_FormatInscription.StdIndent();
                _Label_InfoFormat = G.AjouterLabel(String.Join(" ", ChampsInscription));
                _Label_InfoFormat.StdIndent();


                _CheckBox_InscrireNomTole.OnIsCheck += _Texte_TailleInscription.IsEnable;
                _CheckBox_InscrireNomTole.OnIsCheck += _Texte_FormatInscription.IsEnable;
                _Texte_TailleInscription.IsEnabled = _CheckBox_InscrireNomTole.IsChecked;
                _Texte_FormatInscription.IsEnabled = _CheckBox_InscrireNomTole.IsChecked;



                _CheckBox_OrienterDvp = G.AjouterCheckBox(OrienterDvp);
                _EnumComboBox_OrientationDvp = G.AjouterEnumComboBox<eOrientation, Intitule>(OrientationDvp);
                _EnumComboBox_OrientationDvp.StdIndent();

                _CheckBox_OrienterDvp.OnIsCheck += _EnumComboBox_OrientationDvp.IsEnable;
                _EnumComboBox_OrientationDvp.IsEnabled = _CheckBox_OrienterDvp.IsChecked;

                _CheckBox_FermerPlan = G.AjouterCheckBox(FermerPlan);

                G = _Calque.AjouterGroupe("Fichiers volumineux");
                G.Expanded = true;

                _CheckBox_ConvertirEsquisse = G.AjouterCheckBox(ConvertirEsquisse);

            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private String Quantite()
        {
            CustomPropertyManager PM = App.ModelDoc2.Extension.get_CustomPropertyManager("");

            if (App.ModelDoc2.ePropExiste(PropQuantite.GetValeur<String>()))
            {
                return Math.Max(App.ModelDoc2.eProp(PropQuantite.GetValeur<String>()).eToInteger(), 1).ToString();
            }

            return "1";
        }

        private List<String> ListeMateriaux;

        protected void Rechercher_Materiaux()
        {
            WindowLog.Ecrire("Recherche des materiaux : ");

            ListeMateriaux = App.ModelDoc2.ListeMateriaux(eTypeCorps.Tole);

            foreach (var m in ListeMateriaux)
                WindowLog.Ecrire(" - " + m);

            WindowLog.SautDeLigne();

            _TextListBox_Materiaux.Liste = ListeMateriaux;
            _TextListBox_Materiaux.ToutSelectionner(false);
            _TextComboBox_ForcerMateriau.Liste = ListeMateriaux;
            _TextComboBox_ForcerMateriau.Index = 0;
        }

        protected void RunOkCommand()
        {
            CmdCreerDvp Cmd = new CmdCreerDvp();

            Cmd.MdlBase = App.Sw.ActiveDoc;
            Cmd.ListeMateriaux = _TextListBox_Materiaux.ListSelectedText.Count > 0 ? _TextListBox_Materiaux.ListSelectedText : _TextListBox_Materiaux.Liste;
            Cmd.ForcerMateriau = _CheckBox_ForcerMateriau.IsChecked ? _TextComboBox_ForcerMateriau.Text : null;
            Cmd.Quantite = _Texte_Quantite.Text.eToInteger();
            Cmd.AfficherLignePliage = _CheckBox_AfficherLignePliage.IsChecked;
            Cmd.AfficherNotePliage = _CheckBox_AfficherNotePliage.IsChecked;
            Cmd.InscrireNomTole = _CheckBox_InscrireNomTole.IsChecked;
            Cmd.OrienterDvp = _CheckBox_OrienterDvp.IsChecked;
            Cmd.OrientationDvp = _EnumComboBox_OrientationDvp.Val;
            Cmd.FermerPlan = _CheckBox_FermerPlan.IsChecked;
            Cmd.ConvertirEsquisse = _CheckBox_ConvertirEsquisse.IsChecked;
            Cmd.ComposantsExterne = _CheckBox_ComposantsExterne.IsChecked;
            Cmd.RefFichier = _Texte_RefFichier.Text.Trim();
            Cmd.TailleInscription = _Texte_TailleInscription.Text.eToInteger();
            Cmd.FormatInscription = _Texte_FormatInscription.Text;

            Cmd.Executer();
        }
    }
}
