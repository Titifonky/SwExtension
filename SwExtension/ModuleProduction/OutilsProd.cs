using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ModuleProduction
{
    public static class CONST_PRODUCTION
    {
        public const String DOSSIER_PIECES = "Pieces";
        public const String DOSSIER_PIECES_APERCU = "Apercu";
        public const String FICHIER_NOMENC = "Nomenclature";
        public const String DOSSIER_LASERTOLE = "Laser tole";
        public const String DOSSIER_LASERTUBE = "Laser tube";
        public const String MAX_INDEXDIM = "MAX_INDEXDIM";
        public const String ID_PIECE = "ID_PIECE";
        public const String ID_CONFIG = "ID_CONFIG";
        public const String ID_DOSSIERS = "ID_DOSSIERS";
    }

    public static class OutilsProd
    {

        public static Boolean CreerConfigDepliee(this ModelDoc2 mdl, String NomConfigDepliee, String NomConfigPliee)
        {
            var cfg = mdl.ConfigurationManager.AddConfiguration(NomConfigDepliee, NomConfigDepliee, "", 0, NomConfigPliee, "");
            if (cfg.IsRef())
                return true;

            return false;
        }

        public static void DeplierTole(this PartDoc piece, String nomConfigDepliee)
        {
            var mdl = piece.eModelDoc2();
            var liste = piece.eListeFonctionsDepliee();
            if (liste.Count == 0) return;

            Feature FonctionDepliee = liste[0];
            FonctionDepliee.eModifierEtat(swFeatureSuppressionAction_e.swUnSuppressFeature, nomConfigDepliee);
            FonctionDepliee.eModifierEtat(swFeatureSuppressionAction_e.swUnSuppressDependent, nomConfigDepliee);

            mdl.eEffacerSelection();

            FonctionDepliee.eParcourirSousFonction(
                f =>
                {
                    f.eModifierEtat(swFeatureSuppressionAction_e.swUnSuppressFeature, nomConfigDepliee);

                    if ((f.Name.ToLowerInvariant() == CONSTANTES.LIGNES_DE_PLIAGE.ToLowerInvariant()) ||
                        (f.Name.ToLowerInvariant() == CONSTANTES.CUBE_DE_VISUALISATION.ToLowerInvariant()))
                    {
                        f.eSelect(true);
                    }
                    return false;
                }
                );

            mdl.UnblankSketch();
            mdl.eEffacerSelection();
        }

        public static void PlierTole(this PartDoc piece, String nomConfigPliee)
        {
            var mdl = piece.eModelDoc2();
            var liste = piece.eListeFonctionsDepliee();
            if (liste.Count == 0) return;

            Feature FonctionDepliee = liste[0];

            FonctionDepliee.eModifierEtat(swFeatureSuppressionAction_e.swSuppressFeature, nomConfigPliee);

            mdl.eEffacerSelection();

            FonctionDepliee.eParcourirSousFonction(
                f =>
                {
                    if ((f.Name.ToLowerInvariant() == CONSTANTES.LIGNES_DE_PLIAGE.ToLowerInvariant()) ||
                        (f.Name.ToLowerInvariant() == CONSTANTES.CUBE_DE_VISUALISATION.ToLowerInvariant()))
                    {
                        f.eSelect(true);
                    }
                    return false;
                }
                );

            mdl.BlankSketch();
            mdl.eEffacerSelection();
        }

        /// <summary>
        /// Renvoi la liste unique des modeles et configurations
        /// Modele : ModelDoc2
        ///                   |-Config1 : Nom de la configuration
        ///                   |     |-Nb : quantite de configuration identique dans le modele complet
        ///                   |-Config 2
        ///                   | etc...
        /// </summary>
        /// <param name="mdlBase"></param>
        /// <param name="composantsExterne"></param>
        /// <param name="filtreTypeCorps"></param>
        /// <returns></returns>
        public static SortedDictionary<ModelDoc2, SortedDictionary<String, int>> ListerComposants(this ModelDoc2 mdlBase, Boolean composantsExterne)
        {
            SortedDictionary<ModelDoc2, SortedDictionary<String, int>> dic = new SortedDictionary<ModelDoc2, SortedDictionary<String, int>>(new CompareModelDoc2());

            try
            {
                Action<Component2> Test = delegate (Component2 comp)
                {
                    var cfg = comp.eNomConfiguration();

                    if (comp.IsSuppressed() || comp.ExcludeFromBOM || !cfg.eEstConfigPliee() || comp.TypeDoc() != eTypeDoc.Piece) return;

                    var mdl = comp.eModelDoc2();

                    if (dic.ContainsKey(mdl))
                        if (dic[mdl].ContainsKey(cfg))
                        {
                            dic[mdl][cfg] += 1;
                            return;
                        }

                    if (dic.ContainsKey(mdl))
                        dic[mdl].Add(cfg, 1);
                    else
                    {
                        var lcfg = new SortedDictionary<String, int>(new WindowsStringComparer());
                        lcfg.Add(cfg, 1);
                        dic.Add(mdl, lcfg);
                    }
                };

                if (mdlBase.TypeDoc() == eTypeDoc.Piece)
                    Test(mdlBase.eComposantRacine());
                else
                {
                    mdlBase.eComposantRacine().eRecParcourirComposantBase(
                        Test,
                        // On ne parcourt pas les assemblages exclus
                        c =>
                        {
                            if (c.ExcludeFromBOM)
                                return false;

                            return true;
                        }
                        );
                }
            }
            catch (Exception e) { Log.LogErreur(new Object[] { e }); }

            return dic;
        }

        /// <summary>
        /// Renvoi la liste des modeles avec la liste des configurations, des dossiers
        /// et les quantites de chaque dossier dans l'assemblage
        /// Modele : ModelDoc2
        ///     |-Config1 : Nom de la configuration
        ///     |     |-Dossier1 : Comparaison avec la propriete RefDossier, référence à l'Id de la fonction pour pouvoir le selectionner plus tard
        ///     |     |      |- Nb : quantite de corps identique dans le modele complet
        ///     |     |
        ///     |     |-Dossier2
        ///     |            |- Nb
        ///     |-Config 2
        ///     | etc...
        /// </summary>
        /// <param name="mdlBase"></param>
        /// <param name="composantsExterne"></param>
        /// <param name="filtreDossier"></param>
        /// <returns></returns>
        public static SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<int, int>>> DenombrerDossiers(this ModelDoc2 mdlBase, Boolean composantsExterne, Predicate<Feature> filtreDossier = null, Boolean fermerFichier = false)
        {

            SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<int, int>>> dic = new SortedDictionary<ModelDoc2, SortedDictionary<String, SortedDictionary<int, int>>>(new CompareModelDoc2());
            try
            {
                var ListeComposants = mdlBase.ListerComposants(composantsExterne);

                var ListeDossiers = new Dictionary<String, Dossier>();

                Predicate<Feature> Test = delegate (Feature fDossier)
                {
                    BodyFolder SwDossier = fDossier.GetSpecificFeature2();
                    if (SwDossier.IsRef() && SwDossier.eNbCorps() > 0 && !SwDossier.eEstExclu() && (filtreDossier.IsNull() || filtreDossier(fDossier)))
                        return true;

                    return false;
                };

                foreach (var mdl in ListeComposants.Keys)
                {
                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                    foreach (var t in ListeComposants[mdl])
                    {
                        var cfg = t.Key;
                        var nbCfg = t.Value;
                        mdl.ShowConfiguration2(cfg);
                        mdl.EditRebuild3();
                        var Piece = mdl.ePartDoc();

                        foreach (var fDossier in Piece.eListeDesFonctionsDePiecesSoudees(Test))
                        {
                            BodyFolder SwDossier = fDossier.GetSpecificFeature2();
                            var RefDossier = SwDossier.eProp(CONSTANTES.REF_DOSSIER);

                            if (ListeDossiers.ContainsKey(RefDossier))
                                ListeDossiers[RefDossier].Nb += SwDossier.eNbCorps() * nbCfg;
                            else
                            {
                                var dossier = new Dossier(RefDossier, mdl, cfg, fDossier.GetID());
                                dossier.Nb = SwDossier.eNbCorps() * nbCfg;

                                ListeDossiers.Add(RefDossier, dossier);
                            }
                        }
                    }

                    if (fermerFichier && (mdl.GetPathName() != mdlBase.GetPathName()))
                        App.Sw.CloseDoc(mdl.GetPathName());
                }

                // Conversion d'une liste de dossier
                // en liste de modele
                foreach (var dossier in ListeDossiers.Values)
                {
                    if (dic.ContainsKey(dossier.Mdl))
                    {
                        var lcfg = dic[dossier.Mdl];
                        if (lcfg.ContainsKey(dossier.Config))
                        {
                            var ldossier = lcfg[dossier.Config];
                            ldossier.Add(dossier.Id, dossier.Nb);
                        }
                        else
                        {
                            var ldossier = new SortedDictionary<int, int>();
                            ldossier.Add(dossier.Id, dossier.Nb);
                            lcfg.Add(dossier.Config, ldossier);
                        }
                    }
                    else
                    {
                        var ldossier = new SortedDictionary<int, int>();
                        ldossier.Add(dossier.Id, dossier.Nb);
                        var lcfg = new SortedDictionary<String, SortedDictionary<int, int>>(new WindowsStringComparer());
                        lcfg.Add(dossier.Config, ldossier);
                        dic.Add(dossier.Mdl, lcfg);
                    }
                }
            }
            catch (Exception e) { Log.LogErreur(new Object[] { e }); }

            return dic;
        }

        private class Dossier
        {
            public int Id;
            public String Repere;
            public ModelDoc2 Mdl;
            public String Config;
            public int Nb = 0;

            public Dossier(String repere, ModelDoc2 mdl, String config, int id)
            {
                Repere = repere;
                Mdl = mdl;
                Config = config;
                Id = id;
            }
        }

        private const String ChaineIndice = "ZYXWVUTSRQPONMLKJIHGFEDCBA";

        public static String ChercherIndice(List<String> liste)
        {
            for (int i = 0; i < ChaineIndice.Length; i++)
            {
                if (liste.Any(d => { return d.EndsWith(" Ind " + ChaineIndice[i]) ? true : false; }))
                    return "Ind " + ChaineIndice[Math.Max(0, i - 1)];
            }

            return "Ind " + ChaineIndice.Last();
        }

        public static String CreerDossier(this ModelDoc2 mdl, String dossier)
        {
            var chemin = Path.Combine(mdl.eDossier(), dossier);
            if (!Directory.Exists(chemin))
                Directory.CreateDirectory(chemin);

            return chemin;
        }

        public static String CreerFichierTexte(this ModelDoc2 mdl, String dossier, String fichier)
        {
            var chemin = Path.Combine(mdl.eDossier(), dossier, fichier + ".txt");
            if (!File.Exists(chemin))
                File.WriteAllText(chemin, "", Encoding.GetEncoding(1252));

            return chemin;
        }

        public static int RechercherIndiceDossier(this ModelDoc2 mdl, String dossier)
        {
            int indice = 0;
            String chemin = Path.Combine(mdl.eDossier(), dossier);

            if (Directory.Exists(chemin))
                foreach (var d in Directory.EnumerateDirectories(chemin, "*", SearchOption.TopDirectoryOnly))
                    indice = Math.Max(indice, new DirectoryInfo(d).Name.eToInteger());

            return indice;
        }

        public static String DossierPiece(this ModelDoc2 mdl)
        {
            return Path.Combine(mdl.eDossier(), CONST_PRODUCTION.DOSSIER_PIECES);
        }

        public static String FichierNomenclature(this ModelDoc2 mdl)
        {
            return Path.Combine(mdl.eDossier(), CONST_PRODUCTION.DOSSIER_PIECES, CONST_PRODUCTION.FICHIER_NOMENC + ".txt");
        }

        public static String DossierLaserTole(this ModelDoc2 mdl)
        {
            return Path.Combine(mdl.eDossier(), CONST_PRODUCTION.DOSSIER_LASERTOLE);
        }

        public static String DossierLaserTube(this ModelDoc2 mdl)
        {
            return Path.Combine(mdl.eDossier(), CONST_PRODUCTION.DOSSIER_LASERTUBE);
        }

        public static SortedDictionary<int, Corps> ChargerNomenclature(this ModelDoc2 mdl)
        {
            SortedDictionary<int, Corps> Liste = new SortedDictionary<int, Corps>();

            var chemin = mdl.FichierNomenclature();

            if (File.Exists(chemin))
            {
                using (var sr = new StreamReader(chemin, Encoding.GetEncoding(1252)))
                {
                    // On lit la première ligne contenant l'entête des colonnes
                    String ligne = sr.ReadLine();
                    if (ligne.IsRef())
                    {
                        while ((ligne = sr.ReadLine()) != null)
                        {
                            if (!String.IsNullOrWhiteSpace(ligne))
                            {
                                var c = new Corps(ligne);
                                Liste.Add(c.Repere, c);
                            }
                        }
                    }
                }
            }

            return Liste;
        }

        public static String ExtPiece = eTypeDoc.Piece.GetEnumInfo<ExtFichier>();
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

        public String NomFichier(ModelDoc2 mdlBase)
        {
            return Path.Combine(mdlBase.DossierPiece(), RepereComplet + OutilsProd.ExtPiece);
        }

        public String RepereComplet
        {
            get { return CONSTANTES.PREFIXE_REF_DOSSIER + Repere; }
        }
    }
}
