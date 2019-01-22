using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Linq;

namespace ModuleProduction.ModuleRepererDossier
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Repérer les dossiers"),
        ModuleNom("RepererDossier"),
        ModuleDescription("Repérer les dossiers.")
        ]
    public class PageRepererDossier : BoutonPMPManager
    {
        private Parametre CombinerCorpsIdentiques;
        private Parametre CombinerAvecCampagnePrecedente;
        private Parametre CreerDvp;
        private Parametre TypeCorps;

        private ModelDoc2 MdlBase = null;
        private int _IndiceCampagne = 0;
        private ListeSortedCorps ListeCorps = new ListeSortedCorps();

        private int IndiceCampagne
        {
            get { return _IndiceCampagne; }
            set
            {
                _IndiceCampagne = value;
                _Texte_IndiceCampagne.Text = _IndiceCampagne.ToString();
            }
        }

        public PageRepererDossier()
        {
            try
            {
                CombinerCorpsIdentiques = _Config.AjouterParam("CombinerCorpsIdentiques", true, "Combiner les corps identiques des différents modèles");
                CombinerAvecCampagnePrecedente = _Config.AjouterParam("CombinerAvecCampagnePrecedente", true, "Combiner les corps avec les précédentes campagnes");
                CreerDvp = _Config.AjouterParam("CreerDvp", true, "Creer les configs dvp des tôles");
                TypeCorps = _Config.AjouterParam("TypeCorps", eTypeCorps.Piece, "Type de corps à repérer :");
                TypeCorps.SetValeur(eTypeCorps.Piece);

                MdlBase = App.Sw.ActiveDoc;
                if(MdlBase.ePropExiste(CONST_PRODUCTION.FILTRE_CORPS))
                {
                    var r = (eTypeCorps)Enum.Parse(typeof(eTypeCorps), MdlBase.eGetProp(CONST_PRODUCTION.FILTRE_CORPS));
                    TypeCorps.SetValeur(r);
                }

                OnCalque += Calque;
                OnRunAfterActivation += RechercherIndiceCampagne;
                OnRunOkCommand += RunOkCommand;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private CtrlTextBox _Texte_IndiceCampagne;
        private CtrlCheckBox _CheckBox_MajCampagnePrecedente;
        private CtrlCheckBox _CheckBox_CampagneDepartDecompte;
        private Boolean ReinitCampagneActuelle = false;
        private CtrlCheckBox _CheckBox_ReinitCampagneActuelle;
        private CtrlCheckBox _CheckBox_CombinerCorpsIdentiques;
        private CtrlCheckBox _CheckBox_CombinerAvecCampagnePrecedente;
        private CtrlCheckBox _CheckBox_CreerDvp;
        private CtrlEnumComboBox<eTypeCorps, Intitule> _EnumComboBox_TypeCorps;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Reperage");

                _Texte_IndiceCampagne = G.AjouterTexteBox("Indice de la campagne de repérage :");
                _Texte_IndiceCampagne.LectureSeule = true;

                _CheckBox_CampagneDepartDecompte = G.AjouterCheckBox("Campagne de depart pour les decomptes");
                _CheckBox_ReinitCampagneActuelle = G.AjouterCheckBox("Reinitialiser la campagne actuelle");
                _CheckBox_MajCampagnePrecedente = G.AjouterCheckBox("Mettre à jour la campagne précédente (en cas d'oubli)");
                _CheckBox_MajCampagnePrecedente.StdIndent();

                _CheckBox_ReinitCampagneActuelle.OnCheck += delegate { _CheckBox_MajCampagnePrecedente.IsEnabled = false; _CheckBox_MajCampagnePrecedente.IsChecked = false; };
                _CheckBox_ReinitCampagneActuelle.OnUnCheck += delegate { if (IndiceCampagne > 1) _CheckBox_MajCampagnePrecedente.IsEnabled = true; };

                _CheckBox_MajCampagnePrecedente.OnCheck += delegate { _CheckBox_ReinitCampagneActuelle.IsEnabled = false; _CheckBox_ReinitCampagneActuelle.IsChecked = false; };
                _CheckBox_MajCampagnePrecedente.OnUnCheck += delegate { _CheckBox_ReinitCampagneActuelle.IsEnabled = true; };

                _CheckBox_MajCampagnePrecedente.OnCheck += delegate
                {
                    if (_CheckBox_MajCampagnePrecedente.IsEnabled && (IndiceCampagne > 1))
                        IndiceCampagne -= 1;

                    if (ReinitCampagneActuelle)
                        _CheckBox_ReinitCampagneActuelle.IsEnabled = false;

                    _CheckBox_CampagneDepartDecompte.IsEnabled = false;
                };
                _CheckBox_MajCampagnePrecedente.OnUnCheck += delegate
                {
                    if (_CheckBox_MajCampagnePrecedente.IsEnabled)
                        IndiceCampagne += 1;

                    if (ReinitCampagneActuelle)
                    {
                        _CheckBox_ReinitCampagneActuelle.IsEnabled = true;
                        _CheckBox_ReinitCampagneActuelle.Visible = true;
                    }

                    _CheckBox_CampagneDepartDecompte.IsEnabled = true;
                };

                G = _Calque.AjouterGroupe("Options");
                _CheckBox_CombinerCorpsIdentiques = G.AjouterCheckBox(CombinerCorpsIdentiques);
                _CheckBox_CombinerAvecCampagnePrecedente = G.AjouterCheckBox(CombinerAvecCampagnePrecedente);
                _CheckBox_CombinerAvecCampagnePrecedente.StdIndent();
                _CheckBox_CreerDvp = G.AjouterCheckBox(CreerDvp);
                

                _CheckBox_CombinerCorpsIdentiques.OnUnCheck += _CheckBox_CombinerAvecCampagnePrecedente.UnCheck;
                _CheckBox_CombinerCorpsIdentiques.OnIsCheck += _CheckBox_CombinerAvecCampagnePrecedente.IsEnable;

                _EnumComboBox_TypeCorps = G.AjouterEnumComboBox<eTypeCorps, Intitule>(TypeCorps);
                _EnumComboBox_TypeCorps.FiltrerEnum = eTypeCorps.Tole | eTypeCorps.Barre | eTypeCorps.Piece;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RechercherIndiceCampagne()
        {
            try
            {
                WindowLog.Ecrire("Recherche des éléments existants :");

                // Recherche du dernier indice de la campagne de repérage

                // Création du dossier pièces s'il n'existe pas
                MdlBase.pCreerDossier(CONST_PRODUCTION.DOSSIER_PIECES);
                // Recherche de la nomenclature
                MdlBase.pCreerFichierTexte(CONST_PRODUCTION.DOSSIER_PIECES, CONST_PRODUCTION.FICHIER_NOMENC);

                // Recherche des exports laser, tole ou tube, existant
                var IndiceLaser = Math.Max(MdlBase.pRechercherIndiceDossier(CONST_PRODUCTION.DOSSIER_LASERTOLE),
                                              MdlBase.pRechercherIndiceDossier(CONST_PRODUCTION.DOSSIER_LASERTUBE)
                                              );

                IndiceCampagne = IndiceLaser + 1;

                // Aquisition de la liste des corps déjà repérées
                // et recherche de l'indice nomenclature max
                ListeCorps = MdlBase.pChargerNomenclature();

                WindowLog.EcrireF("{0} pièce(s) sont référencé(s)", ListeCorps.Count);

                // S'il n'y a aucun corps, on désactive les options
                if (ListeCorps.Count == 0)
                {
                    _CheckBox_MajCampagnePrecedente.IsEnabled = false;
                    _CheckBox_MajCampagnePrecedente.Visible = false;

                    _CheckBox_ReinitCampagneActuelle.IsEnabled = false;
                    _CheckBox_ReinitCampagneActuelle.Visible = false;
                }

                // Si aucun corps n'a déjà été comptabilisé pour cette campagne,
                // on ne propose pas la réinitialisation
                foreach (var corps in ListeCorps.Values)
                {
                    if (corps.Campagne.Keys.Max() >= IndiceCampagne)
                    {
                        ReinitCampagneActuelle = true;
                        break;
                    }
                }

                if (!ReinitCampagneActuelle)
                {
                    _CheckBox_ReinitCampagneActuelle.IsEnabled = false;
                    _CheckBox_ReinitCampagneActuelle.Visible = false;
                }


                // Si c'est la première campagne, on désactive des options
                if (IndiceCampagne == 1)
                {
                    _CheckBox_CampagneDepartDecompte.IsEnabled = false;
                    _CheckBox_CampagneDepartDecompte.Visible = false;

                    _CheckBox_MajCampagnePrecedente.IsEnabled = false;
                    _CheckBox_MajCampagnePrecedente.Visible = false;

                    _CheckBox_CombinerAvecCampagnePrecedente.IsEnabled = false;
                    _CheckBox_CombinerAvecCampagnePrecedente.Visible = false;
                }
            }
            catch (Exception e) { this.LogErreur(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdRepererDossier Cmd = new CmdRepererDossier();

            if(_CheckBox_CampagneDepartDecompte.IsChecked)
                ListeCorps.CampagneDepartDecompte = Math.Max(ListeCorps.CampagneDepartDecompte, IndiceCampagne);

            Cmd.MdlBase = App.Sw.ActiveDoc;
            Cmd.IndiceCampagne = IndiceCampagne;
            Cmd.CombinerCorpsIdentiques = _CheckBox_CombinerCorpsIdentiques.IsChecked;
            Cmd.CombinerAvecCampagnePrecedente = _CheckBox_CombinerAvecCampagnePrecedente.IsChecked;
            Cmd.ReinitCampagneActuelle = ReinitCampagneActuelle && _CheckBox_ReinitCampagneActuelle.IsChecked;
            Cmd.CreerDvp = _CheckBox_CreerDvp.IsChecked;
            Cmd.ListeCorps = ListeCorps;
            Cmd.FiltrerCorps = _EnumComboBox_TypeCorps.Val;

            Cmd.Executer();

            MdlBase.ePropAdd(CONST_PRODUCTION.FILTRE_CORPS, _EnumComboBox_TypeCorps.Val);
        }
    }
}