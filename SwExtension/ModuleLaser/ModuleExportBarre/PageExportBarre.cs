using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace ModuleLaser
{
    namespace ModuleExportBarre
    {
        [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
            ModuleTitre("Exporter les barres"),
            ModuleNom("ExportBarre"),
            ModuleDescription("Exporter les barres.")
            ]
        public class PageExportBarre : BoutonPMPManager
        {
            private Parametre PropQuantite;
            private Parametre PrendreEnCompteTole;

            private Parametre ComposantsExterne;
            private Parametre TypeExport;

            private Parametre ExporterBarres;
            private Parametre ListerUsinages;
            private Parametre CreerPdf3D;

            public PageExportBarre()
            {
                try
                {
                    PropQuantite = _Config.AjouterParam("PropQuantite", CONSTANTES.PROPRIETE_QUANTITE, "Propriete \"Quantite\"", "Recherche cette propriete");
                    PrendreEnCompteTole = _Config.AjouterParam("PrendreEnCompteTole", true, "Prendre en compte les tôles");
                    ComposantsExterne = _Config.AjouterParam("ComposantExterne", false, "Exporter les barres externes au dossier du modèle");
                    TypeExport = _Config.AjouterParam("TypeExport", eTypeFichierExport.ParasolidBinary, "Format :");

                    ListerUsinages = _Config.AjouterParam("ListerUsinages", false, "Lister les usinages");
                    ExporterBarres = _Config.AjouterParam("ExporterBarres", false, "Exporter les barres");
                    CreerPdf3D = _Config.AjouterParam("CreerPdf3D", false, "Créer les pdf 3D des barres");

                    OnCalque += Calque;
                    OnRunAfterActivation += Rechercher_Materiaux;
                    OnRunOkCommand += RunOkCommand;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private CtrlTextListBox _TextListBox_Materiaux;
            private CtrlTextBox _Texte_RefFichier;
            private CtrlTextBox _Texte_Quantite;
            private CtrlCheckBox _CheckBox_ComposantsExterne;
            private CtrlCheckBox _CheckBox_PrendreEnCompteTole;
            private CtrlCheckBox _CheckBox_ExporterBarres;
            private CtrlEnumComboBox<eTypeFichierExport, Intitule> _EnumComboBox_TypeExport;
            private CtrlCheckBox _CheckBox_ForcerMateriau;
            private CtrlTextComboBox _TextComboBox_ForcerMateriau;
            private CtrlCheckBox _CheckBox_ListerUsinages;
            private CtrlCheckBox _CheckBox_CreerPdf3D;

            protected void Calque()
            {
                try
                {
                    Groupe G;

                    G = _Calque.AjouterGroupe("Fichier");

                    _Texte_RefFichier = G.AjouterTexteBox("Référence du fichier :", "la référence est ajoutée au début du nom de chaque fichier généré");

                    String Ref = MdlBase.eRefFichier();
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

                    _CheckBox_PrendreEnCompteTole = G.AjouterCheckBox(PrendreEnCompteTole);
                    _CheckBox_PrendreEnCompteTole.OnIsCheck += delegate (Object sender, Boolean value) { Rechercher_Materiaux(); };

                    _CheckBox_ExporterBarres = G.AjouterCheckBox(ExporterBarres);
                    _EnumComboBox_TypeExport = G.AjouterEnumComboBox<eTypeFichierExport, Intitule>(TypeExport);
                    _EnumComboBox_TypeExport.FiltrerEnum = eTypeFichierExport.Parasolid |
                                                            eTypeFichierExport.ParasolidBinary |
                                                            eTypeFichierExport.STEP;

                    _CheckBox_CreerPdf3D = G.AjouterCheckBox(CreerPdf3D);
                    _CheckBox_ExporterBarres.OnIsCheck += _CheckBox_CreerPdf3D.IsEnable;
                    _CheckBox_ExporterBarres.OnIsCheck += _EnumComboBox_TypeExport.IsEnable;
                    _CheckBox_ExporterBarres.ApplyParam();

                    _CheckBox_ListerUsinages = G.AjouterCheckBox(ListerUsinages);
                    

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

            protected void RunOkCommand()
            {
                CmdExportBarre Cmd = new CmdExportBarre();

                Cmd.MdlBase = MdlBase;

                Cmd.ListeMateriaux = _TextListBox_Materiaux.ListSelectedText.Count > 0 ? _TextListBox_Materiaux.ListSelectedText : _TextListBox_Materiaux.Liste;
                Cmd.ForcerMateriau = _CheckBox_ForcerMateriau.IsChecked ? _TextComboBox_ForcerMateriau.Text : null;
                Cmd.Quantite = _Texte_Quantite.Text.eToInteger();
                Cmd.CreerPdf3D = _CheckBox_CreerPdf3D.IsChecked;
                Cmd.TypeExport = _EnumComboBox_TypeExport.Val;
                Cmd.PrendreEnCompteTole = _CheckBox_PrendreEnCompteTole.IsChecked;
                Cmd.ComposantsExterne = _CheckBox_ComposantsExterne.IsChecked;
                Cmd.RefFichier = _Texte_RefFichier.Text;
                Cmd.ExporterBarres = _CheckBox_ExporterBarres.IsChecked;
                Cmd.ListerUsinages = _CheckBox_ListerUsinages.IsChecked;

                Cmd.Executer();

                if (File.Exists(Cmd.CheminNomenclature))
                    System.Diagnostics.Process.Start(Cmd.CheminNomenclature);
            }
        }
    }
}
