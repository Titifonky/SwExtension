using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;

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

            private ModelDoc2 MdlBase = null;
            private int _IndiceCampagne = 1;
            private String DossierPiece = "";
            private String FichierNomenclature = "";
            private List<Corps> ListeCorps = new List<Corps>();

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
                    CombinerCorpsIdentiques = _Config.AjouterParam("CombinerCorpsIdentiques", false, "Combiner les corps identiques des différents modèles");

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
                    _CheckBox_SupprimerReperes.OnCheck += delegate { IndiceCampagne = Math.Max(1, (IndiceCampagne - 1)); };
                    _CheckBox_SupprimerReperes.OnUnCheck += delegate { IndiceCampagne += 1; };
                    _CheckBox_CombinerCorps = G.AjouterCheckBox(CombinerCorpsIdentiques);
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            protected void RechercherIndiceCampagne()
            {
                WindowLog.Ecrire("Recherche des éléments existants :");

                // Recherche du dernier indice de la campagne de repérage

                // Création du dossier pièces s'il n'existe pas
                MdlBase.CreerDossier(OutilsCommun.DossierPieces, out DossierPiece);
                // Recherche de la nomenclature
                MdlBase.CreerFichierTexte(OutilsCommun.DossierPieces, OutilsCommun.FichierNomenclature, out FichierNomenclature);

                // Aquisition de la liste des corps déjà repérées
                int IndiceMin = IndiceCampagne;
                using (var sr = new StreamReader(FichierNomenclature))
                {
                    // On lit la première ligne contenant l'entête des colonnes
                    String ligne = sr.ReadLine();
                    int NbCorps = 0;

                    while ((ligne = sr.ReadLine()) != null)
                    {
                        NbCorps++;
                        if (!String.IsNullOrWhiteSpace(ligne))
                        {
                            var c = new Corps(ligne);
                            IndiceMin = Math.Max(IndiceMin, c.Campagne);
                            ListeCorps.Add(c);
                        }
                    }
                    if (NbCorps > 0)
                        WindowLog.EcrireF("{0} pièce(s) sont référencé(s)", NbCorps);
                }

                // Recherche des exports laser, tole ou tube, existant
                var DossierTole = Path.Combine(MdlBase.eDossier(), OutilsCommun.DossierLaserTole);
                IndiceMin = Math.Max(IndiceMin,
                                     Math.Max(MdlBase.RechercherIndiceDossier(OutilsCommun.DossierLaserTole),
                                              MdlBase.RechercherIndiceDossier(OutilsCommun.DossierLaserTube)
                                              )
                                     );


                if (ListeCorps.Count == 0)
                {
                    _CheckBox_SupprimerReperes.IsChecked = true;
                    _CheckBox_SupprimerReperes.IsEnabled = false;
                }
                else
                {
                    IndiceCampagne = IndiceMin + 1;
                }
            }

            protected void RunOkCommand()
            {
                CmdRepererDossier Cmd = new CmdRepererDossier();

                Cmd.MdlBase = App.Sw.ActiveDoc;
                Cmd.IndiceCampagne = IndiceCampagne;
                Cmd.CombinerCorpsIdentiques = _CheckBox_CombinerCorps.IsChecked;
                Cmd.SupprimerReperes = _CheckBox_SupprimerReperes.IsChecked;
                Cmd.ListeCorpsExistant = ListeCorps;
                Cmd.FichierNomenclature = FichierNomenclature;

                Cmd.Executer();
            }
        }

        public class Corps
        {
            public Body2 SwCorps;
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
            /// SortedDictionary |- Modele
            ///                  |- SortedDictionary |- NomConfig
            ///                                      |- Dictionnary |- Id dossier
            ///                                                     |- Repere
            /// </summary>
            public SortedDictionary<ModelDoc2, SortedDictionary<String, Dictionary<int, int>>> ListeModele = new SortedDictionary<ModelDoc2, SortedDictionary<String, Dictionary<int, int>>>(new CompareModelDoc2());


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

            //public void AjouterModele(ModelDoc2 mdl, String config, int iDDossier)
            //{
            //    if (ListeModele.ContainsKey(mdl))
            //    {
            //        var lCfg = ListeModele[mdl];
            //        if (lCfg.ContainsKey(config))
            //        {
            //            var lDossier = lCfg[config];
            //            if (!lDossier.ContainsKey(iDDossier))
            //                lDossier.Add(iDDossier, Repere);
            //        }
            //        else
            //        {
            //            var lDossier = new Dictionary<int, int>();
            //            lDossier.Add(iDDossier, Repere);
            //            lCfg.Add(config, lDossier);
            //        }
            //    }
            //    else
            //    {
            //        var lDossier = new Dictionary<int, int>();
            //        lDossier.Add(iDDossier, Repere);
            //        var lCfg = new SortedDictionary<String, Dictionary<int, int>>(new WindowsStringComparer());
            //        lCfg.Add(config, lDossier);
            //        ListeModele.Add(mdl, lCfg);
            //    }
            //}

            //public void AjouterModele(Component2 comp, int iDDossier)
            //{
            //    AjouterModele(comp.eModelDoc2(), comp.eNomConfiguration(), iDDossier);
            //}
        }
    }
}
