﻿using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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
        private Parametre CombinerAvecCampagne;
        private Parametre ExporterFichierCorps;

        private ModelDoc2 MdlBase = null;
        private int _IndiceCampagne = 0;
        private String DossierPiece = "";
        private String FichierNomenclature = "";
        private SortedDictionary<int, Corps> ListeCorps = new SortedDictionary<int, Corps>();

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
                CombinerAvecCampagne = _Config.AjouterParam("CombinerAvecPrecedenteCampagne", true, "Combiner les corps avec les précédentes campagnes");
                ExporterFichierCorps = _Config.AjouterParam("ExporterFichierCorps", true, "Exporter les corps dans des fichiers");

                MdlBase = App.Sw.ActiveDoc;
                OnCalque += Calque;
                OnRunAfterActivation += RechercherIndiceCampagne;
                OnRunOkCommand += RunOkCommand;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private CtrlTextBox _Texte_IndiceCampagne;
        private CtrlCheckBox _CheckBox_MajCampagnePrecedente;
        private CtrlCheckBox _CheckBox_ReinitCampagneActuelle;
        private CtrlCheckBox _CheckBox_CombinerCorps;
        private CtrlCheckBox _CheckBox_CombinerAvecCampagne;
        private CtrlCheckBox _CheckBox_ExporterFichierCorps;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Reperage");

                _Texte_IndiceCampagne = G.AjouterTexteBox("Indice de la campagne de repérage :");
                _Texte_IndiceCampagne.LectureSeule = true;

                _CheckBox_ReinitCampagneActuelle = G.AjouterCheckBox("Reinitialiser la campagne actuelle");
                _CheckBox_MajCampagnePrecedente = G.AjouterCheckBox("Mettre à jour la campagne précédente (en cas d'oubli)");

                _CheckBox_ReinitCampagneActuelle.OnCheck += delegate { _CheckBox_MajCampagnePrecedente.IsEnabled = false; _CheckBox_MajCampagnePrecedente.IsChecked = false; };
                _CheckBox_ReinitCampagneActuelle.OnUnCheck += delegate { if (IndiceCampagne > 1) _CheckBox_MajCampagnePrecedente.IsEnabled = true; };

                _CheckBox_MajCampagnePrecedente.OnCheck += delegate { _CheckBox_ReinitCampagneActuelle.IsEnabled = false; _CheckBox_ReinitCampagneActuelle.IsChecked = false; };
                _CheckBox_MajCampagnePrecedente.OnUnCheck += delegate { _CheckBox_ReinitCampagneActuelle.IsEnabled = true; };

                _CheckBox_MajCampagnePrecedente.OnCheck += delegate { if (_CheckBox_MajCampagnePrecedente.IsEnabled && (IndiceCampagne > 1)) IndiceCampagne -= 1; };
                _CheckBox_MajCampagnePrecedente.OnUnCheck += delegate { if (_CheckBox_MajCampagnePrecedente.IsEnabled) IndiceCampagne += 1; };

                G = _Calque.AjouterGroupe("Options");
                _CheckBox_ExporterFichierCorps = G.AjouterCheckBox(ExporterFichierCorps);
                _CheckBox_CombinerCorps = G.AjouterCheckBox(CombinerCorpsIdentiques);
                _CheckBox_CombinerAvecCampagne = G.AjouterCheckBox(CombinerAvecCampagne);
                _CheckBox_CombinerAvecCampagne.Indent = 1;

                _CheckBox_ExporterFichierCorps.OnUnCheck += _CheckBox_CombinerCorps.UnCheck;
                _CheckBox_ExporterFichierCorps.OnIsCheck += _CheckBox_CombinerCorps.IsEnable;
                _CheckBox_ExporterFichierCorps.OnUnCheck += _CheckBox_CombinerAvecCampagne.UnCheck;
                _CheckBox_ExporterFichierCorps.OnIsCheck += _CheckBox_CombinerAvecCampagne.IsEnable;

                _CheckBox_CombinerCorps.OnUnCheck += _CheckBox_CombinerAvecCampagne.UnCheck;
                _CheckBox_CombinerCorps.OnIsCheck += _CheckBox_CombinerAvecCampagne.IsEnable;

                _CheckBox_ExporterFichierCorps.ApplyParam();
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
                MdlBase.CreerDossier(CONST_PRODUCTION.DOSSIER_PIECES, out DossierPiece);
                // Recherche de la nomenclature
                MdlBase.CreerFichierTexte(CONST_PRODUCTION.DOSSIER_PIECES, CONST_PRODUCTION.FICHIER_NOMENC, out FichierNomenclature);

                // Recherche des exports laser, tole ou tube, existant
                var IndiceLaser = Math.Max(MdlBase.RechercherIndiceDossier(CONST_PRODUCTION.DOSSIER_LASERTOLE),
                                              MdlBase.RechercherIndiceDossier(CONST_PRODUCTION.DOSSIER_LASERTUBE)
                                              );

                IndiceCampagne = IndiceLaser + 1;

                // Aquisition de la liste des corps déjà repérées
                // et recherche de l'indice nomenclature max
                using (var sr = new StreamReader(FichierNomenclature, Encoding.GetEncoding(1252)))
                {
                    // On lit la première ligne contenant l'entête des colonnes
                    String ligne = sr.ReadLine();
                    if (ligne.IsRef())
                    {
                        int NbCorps = 0;

                        while ((ligne = sr.ReadLine()) != null)
                        {
                            NbCorps++;
                            if (!String.IsNullOrWhiteSpace(ligne))
                            {
                                var c = new Corps(ligne);
                                ListeCorps.Add(c.Repere, c);
                            }
                        }
                        if (NbCorps > 0)
                            WindowLog.EcrireF("{0} pièce(s) sont référencé(s)", NbCorps);
                    }
                }

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
                Boolean ReinitCampagneActuelle = false;
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
                    _CheckBox_MajCampagnePrecedente.IsEnabled = false;
                    _CheckBox_MajCampagnePrecedente.Visible = false;

                    _CheckBox_CombinerAvecCampagne.IsChecked = false;
                    _CheckBox_CombinerAvecCampagne.IsEnabled = false;
                    _CheckBox_CombinerAvecCampagne.Visible = false;
                }
            }
            catch (Exception e) { this.LogErreur(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdRepererDossier Cmd = new CmdRepererDossier();

            Cmd.MdlBase = App.Sw.ActiveDoc;
            Cmd.IndiceCampagne = IndiceCampagne;
            Cmd.CombinerCorpsIdentiques = _CheckBox_CombinerCorps.IsChecked;
            Cmd.CombinerAvecCampagne = _CheckBox_CombinerAvecCampagne.IsChecked;
            Cmd.ReinitCampagneActuelle = _CheckBox_ReinitCampagneActuelle.IsChecked;
            Cmd.MajCampagnePrecedente = _CheckBox_MajCampagnePrecedente.IsChecked;

            Cmd.ExporterFichierCorps = _CheckBox_ExporterFichierCorps.IsChecked;
            Cmd.ListeCorpsExistant = ListeCorps;
            Cmd.FichierNomenclature = FichierNomenclature;

            Cmd.Executer();
        }
    }

    public class Corps
    {
        public Body2 SwCorps = null;
        public SortedDictionary<int, int> Campagne = new SortedDictionary<int, int>();
        public int Repere;
        public eTypeCorps TypeCorps;
        /// <summary>
        /// Epaisseur de la tôle ou section
        /// </summary>
        public String Dimension;
        public String Materiau;
        public ModelDoc2 Modele = null;
        private long _TailleFichier = long.MaxValue;
        public String NomConfig = "";
        public int IdDossier = -1;
        public String NomCorps = "";

        public static String Entete(int indiceCampagne)
        {
            String entete = String.Format("{0}\t{1}\t{2}\t{3}", "Repere", "Type", "Dimension", "Materiau");
            for (int i = 0; i < indiceCampagne; i++)
                entete += String.Format("\t{0}", i + 1);

            return entete;
        }

        public override string ToString()
        {
            String Ligne = String.Format("{0}\t{1}\t{2}\t{3}", Repere, TypeCorps, Dimension, Materiau);

            for (int i = 0; i < Campagne.Keys.Max(); i++)
            {
                int nb = 0;
                if (Campagne.ContainsKey(i + 1))
                    nb = Campagne[i + 1];

                Ligne += String.Format("\t{0}", nb);
            }

            return Ligne;
        }

        public void InitCampagne(int indiceCampagne)
        {
            if (Campagne.ContainsKey(indiceCampagne))
                Campagne[indiceCampagne] = 0;
            else
                Campagne.Add(indiceCampagne, 0);
        }

        public void InitDimension(BodyFolder dossier, Body2 corps)
        {
            if (TypeCorps == eTypeCorps.Tole)
                Dimension = corps.eEpaisseurCorpsOuDossier(dossier).ToString();
            else
                Dimension = dossier.eProfilDossier();
        }

        public Corps(Body2 swCorps, eTypeCorps typeCorps, String materiau)
        {
            SwCorps = swCorps;
            TypeCorps = typeCorps;
            Materiau = materiau;
        }

        public Corps(eTypeCorps typeCorps, String materiau, String dimension, int campagne, int repere)
        {
            TypeCorps = typeCorps;
            Materiau = materiau;
            Dimension = dimension;
            Campagne.Add(campagne, 0);
            Repere = repere;
        }

        public Corps(String ligne)
        {
            var tab = ligne.Split(new char[] { '\t' });
            Repere = tab[0].eToInteger();
            TypeCorps = (eTypeCorps)Enum.Parse(typeof(eTypeCorps), tab[1]);
            Dimension = tab[2];
            Materiau = tab[3];
            int cp = 1;
            Campagne = new SortedDictionary<int, int>();
            for (int i = 4; i < tab.Length; i++)
                Campagne.Add(cp++, tab[i].eToInteger());
        }

        public void AjouterModele(ModelDoc2 mdl, String config, int idDossier, String nomCorps)
        {
            long t = new FileInfo(mdl.GetPathName()).Length;
            if (t < _TailleFichier)
            {
                _TailleFichier = t;
                Modele = mdl;
                NomConfig = config;
                IdDossier = idDossier;
                NomCorps = nomCorps;
            }
        }
    }
}