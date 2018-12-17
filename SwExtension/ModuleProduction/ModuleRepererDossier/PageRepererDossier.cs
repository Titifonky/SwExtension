using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ModuleProduction
{
    namespace ModuleRepererDossier
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
            private SortedDictionary<int, SortedDictionary<int, Corps>> ListeCorps = new SortedDictionary<int, SortedDictionary<int, Corps>>();

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
            private CtrlCheckBox _CheckBox_CombinerCorps;
            private CtrlCheckBox _CheckBox_CombinerAvecCampagne;
            private CtrlCheckBox _CheckBox_ExporterFichierCorps;
            private CtrlCheckBox _CheckBox_SupprimerReperes;

            protected void Calque()
            {
                try
                {
                    Groupe G;

                    G = _Calque.AjouterGroupe("Options");

                    _Texte_IndiceCampagne = G.AjouterTexteBox("Indice de la campagne de repérage :");
                    _Texte_IndiceCampagne.LectureSeule = true;
                    _CheckBox_SupprimerReperes = G.AjouterCheckBox("Supprimer les repères de la précédente campagne");
                    _CheckBox_ExporterFichierCorps = G.AjouterCheckBox(ExporterFichierCorps);
                    _CheckBox_CombinerCorps = G.AjouterCheckBox(CombinerCorpsIdentiques);
                    _CheckBox_CombinerAvecCampagne = G.AjouterCheckBox(CombinerAvecCampagne);

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

                    // Aquisition de la liste des corps déjà repérées
                    // et recherche de l'indice nomenclature max
                    int IndiceNomenclature = 1;
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
                                    IndiceNomenclature = Math.Max(IndiceNomenclature, c.Campagne);
                                    if (!ListeCorps.ContainsKey(c.Campagne))
                                        ListeCorps.Add(c.Campagne, new SortedDictionary<int, Corps>());

                                    ListeCorps[c.Campagne].Add(c.Repere, c);
                                }
                            }
                            if (NbCorps > 0)
                                WindowLog.EcrireF("{0} pièce(s) sont référencé(s)", NbCorps);
                        }
                    }

                    // Recherche des exports laser, tole ou tube, existant
                    var IndiceLaser = Math.Max(MdlBase.RechercherIndiceDossier(CONST_PRODUCTION.DOSSIER_LASERTOLE),
                                                  MdlBase.RechercherIndiceDossier(CONST_PRODUCTION.DOSSIER_LASERTUBE)
                                                  );

                    if (ListeCorps.Count == 0)
                    {
                        _CheckBox_SupprimerReperes.IsChecked = true;
                        _CheckBox_SupprimerReperes.IsEnabled = false;
                    }

                    if (IndiceLaser >= IndiceNomenclature)
                        IndiceCampagne = IndiceLaser + 1;
                    else
                        IndiceCampagne = IndiceNomenclature;

                    if (IndiceCampagne == 1)
                        _CheckBox_CombinerAvecCampagne.Visible = false;
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
                Cmd.SupprimerReperes = _CheckBox_SupprimerReperes.IsChecked;
                Cmd.ExporterFichierCorps = _CheckBox_ExporterFichierCorps.IsChecked;
                Cmd.ListeCampagnes = ListeCorps;
                Cmd.FichierNomenclature = FichierNomenclature;

                Cmd.Executer();
            }
        }

        public class Corps
        {
            public Body2 SwCorps = null;
            public int Campagne;
            public int Repere;
            public int Nb = 0;
            public eTypeCorps TypeCorps;
            /// <summary>
            /// Epaisseur de la tôle ou section
            /// </summary>
            public String Dimension;
            public String Materiau;

            /// <summary>
            /// SortedDictionary
            ///     |- Modele / SortedDictionary
            ///         |- NomConfig / SortedDictionary
            ///             |- Id dossier / NomCorps
            /// </summary>
            public SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<int, String>>> ListeModele = new SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<int, String>>>(new CompareModelDoc2());


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
                Campagne = campagne;
                Repere = repere;
            }

            public Corps(String ligne)
            {
                var tab = ligne.Split(new char[] { '\t' });
                Campagne = tab[0].eToInteger();
                Repere = tab[1].eToInteger();
                Nb = tab[2].eToInteger();
                TypeCorps = (eTypeCorps)Enum.Parse(typeof(eTypeCorps), tab[3]);
                Dimension = tab[4];
                Materiau = tab[5];
            }

            public void AjouterModele(ModelDoc2 mdl, String config, int iDDossier, String nomCorps)
            {
                if (ListeModele.ContainsKey(mdl))
                {
                    var lCfg = ListeModele[mdl];
                    if (lCfg.ContainsKey(config))
                    {
                        var lDossier = lCfg[config];
                        if (!lDossier.ContainsKey(iDDossier))
                            lDossier.Add(iDDossier, nomCorps);
                    }
                    else
                    {
                        var lDossier = new SortedDictionary<int, String>();
                        lDossier.Add(iDDossier, nomCorps);
                        lCfg.Add(config, lDossier);
                    }
                }
                else
                {
                    var lDossier = new SortedDictionary<int, String>();
                    lDossier.Add(iDDossier, nomCorps);
                    var lCfg = new SortedDictionary<String, SortedDictionary<int, String>>(new WindowsStringComparer());
                    lCfg.Add(config, lDossier);
                    ListeModele.Add(mdl, lCfg);
                }
            }

            public void AjouterModele(Component2 comp, int iDDossier, String nomCorps)
            {
                AjouterModele(comp.eModelDoc2(), comp.eNomConfiguration(), iDDossier, nomCorps);
            }
        }
    }
}
