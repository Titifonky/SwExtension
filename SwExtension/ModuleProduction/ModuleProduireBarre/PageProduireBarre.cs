using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ModuleProduction.ModuleProduireBarre
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Exporter les barres"),
        ModuleNom("ProduireBarre"),
        ModuleDescription("Exporter les barres.")
        ]
    public class PageProduireBarre : BoutonPMPManager
    {
        private Parametre TypeExport;
        private Parametre ExporterBarres;
        private Parametre ListerUsinages;
        private Parametre CreerPdf3D;

        public PageProduireBarre()
        {
            try
            {
                TypeExport = _Config.AjouterParam("TypeExport", eTypeFichierExport.ParasolidBinary, "Format :");

                ListerUsinages = _Config.AjouterParam("ListerUsinages", false, "Lister les usinages");
                ExporterBarres = _Config.AjouterParam("ExporterBarres", false, "Exporter les barres");
                CreerPdf3D = _Config.AjouterParam("CreerPdf3D", false, "Créer les pdf 3D des barres");

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
        private CtrlCheckBox _CheckBox_MettreAjourCampagne;
        private CtrlTextBox _Texte_Quantite;
        private CtrlTextListBox _TextListBox_Materiaux;
        private CtrlTextListBox _TextListBox_Ep;
        private CtrlCheckBox _CheckBox_ExporterBarres;
        private CtrlEnumComboBox<eTypeFichierExport, Intitule> _EnumComboBox_TypeExport;
        private CtrlCheckBox _CheckBox_ListerUsinages;
        private CtrlCheckBox _CheckBox_CreerPdf3D;

        protected void Calque()
        {
            try
            {
                MdlBase = App.ModelDoc2;

                Groupe G;

                G = _Calque.AjouterGroupe("Fichier");

                _Texte_RefFichier = G.AjouterTexteBox("Référence du fichier :", "la référence est ajoutée au début du nom de chaque fichier généré");

                _Texte_RefFichier.Text = MdlBase.eRefFichierComplet();
                _Texte_RefFichier.LectureSeule = false;

                // S'il n'y a pas de reference, on met le texte en rouge
                if (String.IsNullOrWhiteSpace(_Texte_RefFichier.Text))
                    _Texte_RefFichier.BackgroundColor(Color.Red, true);

                _TextBox_Campagne = G.AjouterTexteBox("Campagne :", "");
                _TextBox_Campagne.LectureSeule = true;

                _CheckBox_MettreAjourCampagne = G.AjouterCheckBox("Mettre à jour la campagne");

                G = _Calque.AjouterGroupe("Quantité :");

                _Texte_Quantite = G.AjouterTexteBox("Multiplier par quantité :", "Multiplier les quantités par");
                _Texte_Quantite.Text = MdlBase.Quantite();
                _Texte_Quantite.ValiderTexte += ValiderTextIsInteger;

                G = _Calque.AjouterGroupe("Materiaux :");

                _TextListBox_Materiaux = G.AjouterTextListBox();
                _TextListBox_Materiaux.TouteHauteur = true;
                _TextListBox_Materiaux.Height = 50;
                _TextListBox_Materiaux.SelectionMultiple = true;

                G = _Calque.AjouterGroupe("Profil :");

                _TextListBox_Ep = G.AjouterTextListBox();
                _TextListBox_Ep.TouteHauteur = true;
                _TextListBox_Ep.Height = 50;
                _TextListBox_Ep.SelectionMultiple = true;

                G = _Calque.AjouterGroupe("Options");

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

        private int Campagne;
        private List<String> ListeMateriaux;
        private List<String> ListeProfil;
        private ListeSortedCorps ListeCorps;

        protected void Rechercher_Infos()
        {
            try
            {
                WindowLog.Ecrire("Recherche des materiaux et epaisseurs ");

                ListeCorps = MdlBase.pChargerNomenclature(eTypeCorps.Barre);
                ListeMateriaux = new List<String>();
                ListeProfil = new List<String>();
                Campagne = 1;

                foreach (var corps in ListeCorps.Values)
                {
                    Campagne = Math.Max(Campagne, corps.Campagne.Keys.Max());

                    ListeMateriaux.AddIfNotExist(corps.Materiau);
                    ListeProfil.AddIfNotExist(corps.Dimension);
                }

                WindowLog.SautDeLigne();

                ListeMateriaux.Sort(new WindowsStringComparer());
                ListeProfil.Sort(new WindowsStringComparer());

                _TextBox_Campagne.Text = Campagne.ToString();
                _TextListBox_Materiaux.Liste = ListeMateriaux;
                _TextListBox_Materiaux.ToutSelectionner(false);

                _TextListBox_Ep.Liste = ListeProfil;
                _TextListBox_Ep.ToutSelectionner(false);

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
            CmdProduireBarre Cmd = new CmdProduireBarre();

            Cmd.MdlBase = App.Sw.ActiveDoc;

            Cmd.ListeMateriaux = _TextListBox_Materiaux.ListSelectedText.Count > 0 ? _TextListBox_Materiaux.ListSelectedText : _TextListBox_Materiaux.Liste;
            Cmd.Quantite = _Texte_Quantite.Text.eToInteger();
            Cmd.CreerPdf3D = _CheckBox_CreerPdf3D.IsChecked;
            Cmd.TypeExport = _EnumComboBox_TypeExport.Val;
            Cmd.RefFichier = _Texte_RefFichier.Text;
            Cmd.ExporterBarres = _CheckBox_ExporterBarres.IsChecked;
            Cmd.ListerUsinages = _CheckBox_ListerUsinages.IsChecked;

            Cmd.Executer();
        }
    }
}
