using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using System.Text.RegularExpressions;

namespace ModuleProduction
{
    namespace ModuleRepererDossier
    {
        public class CmdRepererDossier : Cmd
        {
            public ModelDoc2 MdlBase = null;
            public int IndiceCampagne = 0;

            public Boolean SupprimerReperes = false;
            public Boolean CombinerCorpsIdentiques = false;
            public Boolean CombinerAvecCampagne = false;
            public Boolean ExporterFichierCorps = false;
            public String FichierNomenclature = "";
            public String CheminDossierPieces = "";
            public SortedDictionary<int, SortedDictionary<int, Corps>> ListeCampagnes = new SortedDictionary<int, SortedDictionary<int, Corps>>();
            public SortedDictionary<int, String> ListeCorpsCharge = new SortedDictionary<int, String>();

            private int _GenRepereDossier = 0;
            private int GenRepereDossier { get { return ++_GenRepereDossier; } }

            /// <summary>
            /// Pour pouvoir obtenir une référence unique pour chaque dossier de corps, identiques
            /// dans l'assemblage, on passe par la création d'une propriété dans chaque dossier.
            /// Cette propriété est liée à une cote dans une esquisse dont la valeur est égale au repère.
            /// Suivant la configuration, la valeur de la cote peut changer et donc le repère du dossier.
            /// C'est le seul moyen pour avoir un lien entre les dossiers et la configuration du modèle.
            /// Les propriétés des dossiers ne sont pas configurables.
            /// </summary>
            protected override void Command()
            {
                try
                {
                    CheminDossierPieces = Path.Combine(MdlBase.eDossier(), CONST_PRODUCTION.DOSSIER_PIECES);

                    if (SupprimerReperes)
                    {
                        // Si aucun corps n'a déjà été repéré, on reinitialise tout
                        if (ListeCampagnes.Count == 0)
                        {
                            NettoyerModele();

                            // On supprime tout les fichiers
                            foreach (FileInfo file in new DirectoryInfo(CheminDossierPieces).GetFiles())
                                file.Delete();
                        }
                        else
                        {
                            // On supprime les repères de la campagne actuelle
                            String ext = eTypeDoc.Piece.GetEnumInfo<ExtFichier>();

                            // On recherche les repères appartenant aux campagnes précédentes
                            // pour ne pas supprimer les fichiers
                            HashSet<int> FichierAsauvegarder = new HashSet<int>();
                            foreach (var listecorps in ListeCampagnes)
                            {
                                if (listecorps.Key != IndiceCampagne)
                                    foreach (var repere in listecorps.Value.Keys)
                                        FichierAsauvegarder.AddIfNotExist(repere);
                            }

                            // On nettoie les fichiers précedement crées
                            if (ListeCampagnes.ContainsKey(IndiceCampagne))
                            {
                                foreach (var repere in ListeCampagnes[IndiceCampagne].Keys)
                                {
                                    if (FichierAsauvegarder.Contains(repere)) continue;

                                    String fichier = Path.Combine(CheminDossierPieces, CONSTANTES.PREFIXE_REF_DOSSIER + repere + ext);
                                    if (File.Exists(fichier))
                                        File.Delete(fichier);
                                }

                                ListeCampagnes.Remove(IndiceCampagne);
                            }
                        }
                    }

                    // On charge les corps existant à partir des fichiers
                    if (CombinerCorpsIdentiques && CombinerAvecCampagne && (ListeCampagnes.Count > 0))
                    {
                        WindowLog.SautDeLigne();
                        WindowLog.Ecrire("Chargement des corps existants :");
                        foreach (FileInfo file in new DirectoryInfo(CheminDossierPieces).GetFiles())
                        {
                            int rep = Path.GetFileNameWithoutExtension(file.Name).Replace(CONSTANTES.PREFIXE_REF_DOSSIER, "").eToInteger();

                            foreach (var listecorps in ListeCampagnes.Values)
                            {
                                if (listecorps.ContainsKey(rep))
                                {
                                    WindowLog.EcrireF("- {0}", Path.GetFileNameWithoutExtension(file.Name));
                                    ListeCorpsCharge.Add(rep, file.FullName);
                                    ModelDoc2 mdl = Sw.eOuvrir(file.FullName);
                                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                                    var Piece = mdl.ePartDoc();
                                    listecorps[rep].SwCorps = Piece.ePremierCorps();
                                    break;
                                }
                            }
                        }
                    }

                    MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                    // On cree la liste pour cette campagne
                    // Si elle existe, on reinitialise les quantité à 0
                    if (!ListeCampagnes.ContainsKey(IndiceCampagne))
                        ListeCampagnes.Add(IndiceCampagne, new SortedDictionary<int, Corps>());
                    else
                        foreach (var corps in ListeCampagnes[IndiceCampagne].Values)
                            corps.Nb = 0;

                    // On recherche l'indice de repere max
                    foreach (var listecorps in ListeCampagnes.Values)
                        foreach (var repere in listecorps.Keys)
                            _GenRepereDossier = Math.Max(_GenRepereDossier, repere);

                    // On liste les composants
                    var lst = MdlBase.ListerComposants(false);

                    // On boucle sur les modeles
                    foreach (var mdl in lst.Keys)
                    {
                        mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                        // On crée l'esquisse pour le reperage des dossiers
                        EsquisseRepere(mdl);

                        // Si le modele est a repérer complètement
                        Boolean InitModele = true;
                        // On recherche l'index de la dimension maximum
                        int IndexDimension = 0;

                        // Les données sont stockées dans des propriétés du modèle
                        // Le nom du modèle est stocké dans une propriété, si le modèle est copié
                        // la propriété n'est plus valable, on force le repérage
                        // On récupère également le dernier indice de la dimension utilisée
                        if (mdl.ePropExiste(CONST_PRODUCTION.ID_PIECE) && (mdl.eProp(CONST_PRODUCTION.ID_PIECE) == mdl.eNomSansExt()))
                        {
                            InitModele = false;
                            if (mdl.ePropExiste(CONST_PRODUCTION.MAX_INDEXDIM))
                                IndexDimension = mdl.eProp(CONST_PRODUCTION.MAX_INDEXDIM).eToInteger();
                        }

                        foreach (var nomCfg in lst[mdl].Keys)
                        {
                            mdl.ShowConfiguration2(nomCfg);
                            mdl.EditRebuild3();
                            WindowLog.SautDeLigne();
                            WindowLog.EcrireF("{0} \"{1}\"", mdl.eNomSansExt(), nomCfg);

                            HashSet<int> ListIdDossiers = new HashSet<int>();

                            Boolean InitConfig = true;

                            int IdCfg = mdl.GetConfigurationByName(nomCfg).GetID();

                            // Idem modèle, on stock l'id de la config dans une propriété.
                            // Si une nouvelle config est crée, la valeur de cette propriété devient caduc,
                            // on repère alors les dossiers
                            // On en profite pour récupérer la liste des ids de dossiers déjà traité dans les précédentes
                            // campagne de repérage
                            if (!InitModele && mdl.ePropExiste(CONST_PRODUCTION.ID_CONFIG, nomCfg) && (mdl.eProp(CONST_PRODUCTION.ID_CONFIG, nomCfg) == IdCfg.ToString()))
                            {
                                InitConfig = false;
                                if (mdl.ePropExiste(CONST_PRODUCTION.ID_DOSSIERS, nomCfg))
                                {
                                    var tab = mdl.eProp(CONST_PRODUCTION.ID_DOSSIERS, nomCfg).Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (var id in tab)
                                        ListIdDossiers.Add(id.eToInteger());
                                }
                            }

                            var piece = mdl.ePartDoc();
                            var ListeDossier = piece.eListeDesFonctionsDePiecesSoudees(
                                swD =>
                                {
                                    BodyFolder Dossier = swD.GetSpecificFeature2();

                                    // Si le dossier est la racine d'un sous-ensemble soudé, il n'y a rien dedans
                                    if (Dossier.IsRef() && (Dossier.eNbCorps() > 0) &&
                                        (eTypeCorps.Barre | eTypeCorps.Tole).HasFlag(Dossier.eTypeDeDossier()))
                                        return true;

                                    return false;
                                }
                                );

                            var NbConfig = lst[mdl][nomCfg];

                            foreach (var fDossier in ListeDossier)
                            {
                                BodyFolder Dossier = fDossier.GetSpecificFeature2();
                                int IdDossier = fDossier.GetID();

                                WindowLog.EcrireF("     {0}", fDossier.Name);

                                Boolean NouveauDossier = false;
                                Dimension param = GetParam(mdl, fDossier, nomCfg, ref IndexDimension, out NouveauDossier);

                                var SwCorps = Dossier.ePremierCorps();
                                var NomCorps = SwCorps.Name;
                                var MateriauCorps = SwCorps.eGetMateriauCorpsOuPiece(piece, nomCfg);
                                eTypeCorps TypeCorps = Dossier.eTypeDeDossier();
                                var nbCorps = Dossier.eNbCorps() * NbConfig;

                                Boolean Combiner = false;
                                int Repere = -1;
                                if (CombinerCorpsIdentiques)
                                {
                                    // On recherche s'il existe des corps identiques
                                    // Si oui, on applique le même repère au parametre

                                    foreach (var campagne in ListeCampagnes.Keys)
                                    {
                                        if (CombinerAvecCampagne || (campagne == IndiceCampagne))
                                        {
                                            var listecorps = ListeCampagnes[campagne];
                                            foreach (var CorpsTest in listecorps.Values)
                                            {
                                                if (CorpsTest.SwCorps.IsNull() || (MateriauCorps != CorpsTest.Materiau) || (TypeCorps != CorpsTest.TypeCorps)) continue;

                                                if (SwCorps.eEstSemblable(CorpsTest.SwCorps))
                                                {
                                                    Repere = CorpsTest.Repere;
                                                    SetRepere(param, CorpsTest.Repere, nomCfg);
                                                    Combiner = true;
                                                    break;
                                                }
                                            }
                                            if (Combiner) break;
                                        }
                                    }
                                    
                                }

                                Corps corps = null;
                                if (!Combiner)
                                {
                                    Repere = GetRepere(param, nomCfg);

                                    // Création d'un nouveau repère si
                                    if (InitConfig ||
                                        NouveauDossier ||
                                        !ListIdDossiers.Contains(IdDossier) ||
                                        !ListeCampagnes[IndiceCampagne].ContainsKey(Repere))
                                    {
                                        Repere = GenRepereDossier;
                                        SetRepere(param, Repere, nomCfg);
                                    }

                                    if (NouveauDossier || !ListeCampagnes[IndiceCampagne].ContainsKey(Repere))
                                    {
                                        corps = new Corps(SwCorps, TypeCorps, MateriauCorps);
                                        corps.Campagne = IndiceCampagne;
                                        ListeCampagnes[IndiceCampagne].Add(Repere, corps);
                                    }
                                    else
                                        corps = ListeCampagnes[IndiceCampagne][Repere];

                                    corps.Nb = nbCorps;

                                    corps.Repere = Repere;

                                    if (corps.TypeCorps == eTypeCorps.Tole)
                                        corps.Dimension = SwCorps.eEpaisseurCorpsOuDossier(Dossier).ToString();
                                    else
                                        corps.Dimension = Dossier.eProfilDossier();

                                    corps.AjouterModele(mdl, nomCfg, IdDossier, NomCorps);
                                }
                                else
                                {
                                    if (ListeCampagnes[IndiceCampagne].ContainsKey(Repere))
                                        corps = ListeCampagnes[IndiceCampagne][Repere];
                                    else
                                    {
                                        corps = new Corps(SwCorps, TypeCorps, MateriauCorps);
                                        ListeCampagnes[IndiceCampagne].Add(Repere, corps);

                                        corps.Campagne = IndiceCampagne;
                                        corps.Repere = Repere;

                                        if (corps.TypeCorps == eTypeCorps.Tole)
                                            corps.Dimension = SwCorps.eEpaisseurCorpsOuDossier(Dossier).ToString();
                                        else
                                            corps.Dimension = Dossier.eProfilDossier();
                                    }

                                    corps.Nb += nbCorps;
                                    
                                }
                                corps.AjouterModele(mdl, nomCfg, IdDossier, NomCorps);
                                ListIdDossiers.Add(IdDossier);
                            }
                            mdl.ePropAdd(CONST_PRODUCTION.ID_CONFIG, IdCfg, nomCfg);
                            mdl.ePropAdd(CONST_PRODUCTION.ID_DOSSIERS, String.Join(" ", ListIdDossiers), nomCfg);
                        }
                        mdl.ePropAdd(CONST_PRODUCTION.ID_PIECE, mdl.eNomSansExt());
                        mdl.ePropAdd(CONST_PRODUCTION.MAX_INDEXDIM, IndexDimension);
                        mdl.eSauver();

                        if (mdl.GetPathName() != MdlBase.GetPathName())
                            App.Sw.CloseDoc(mdl.GetPathName());
                    }

                    MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                    MdlBase.EditRebuild3();
                    MdlBase.eSauver();

                    ////////////////////////////////// FIN DU REPERAGE ////////////////////////////////////////////////////

                    // On fermer les fichiers chargé
                    foreach (var f in ListeCorpsCharge.Values)
                        App.Sw.CloseDoc(f);

                    ////////////////////////////////// EXPORTER LES CORPS /////////////////////////////////////////////////

                    if (ExporterFichierCorps)
                    {
                        WindowLog.SautDeLigne();
                        WindowLog.Ecrire("Export des corps :");

                        String ext = eTypeDoc.Piece.GetEnumInfo<ExtFichier>();

                        foreach (var corps in ListeCampagnes[IndiceCampagne].Values)
                        {
                            if (corps.ListeModele.Count == 0) continue;

                            var cheminFichier = Path.Combine(CheminDossierPieces, CONSTANTES.PREFIXE_REF_DOSSIER + corps.Repere + ext);
                            if (File.Exists(cheminFichier)) continue;

                            ModelDoc2 mdl = null; String nomCfg = ""; int idDossier = -1; String nomCorps = "";

                            // On recherche le fichier le plus petit à lier
                            long taille = 0;
                            foreach (var m in corps.ListeModele.Keys)
                            {
                                if (new FileInfo(m.GetPathName()).Length > taille)
                                {
                                    var kv1 = corps.ListeModele[m].First();
                                    var kv2 = kv1.Value.First();
                                    mdl = m; nomCfg = kv1.Key; idDossier = kv2.Key; nomCorps = kv2.Value;
                                }
                            }

                            

                            mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                            mdl.ShowConfiguration2(nomCfg);
                            mdl.EditRebuild3();

                            var mdlFichier = Sw.eCreerDocument(CheminDossierPieces, CONSTANTES.PREFIXE_REF_DOSSIER + corps.Repere, eTypeDoc.Piece);
                            mdlFichier.eActiver();
                            var Piece = mdlFichier.ePartDoc();
                            var fonc = Piece.InsertPart3(mdl.GetPathName(), 90241, nomCfg);
                            if (fonc.IsRef())
                            {
                                WindowLog.EcrireF("- {0} exporté", CONSTANTES.PREFIXE_REF_DOSSIER + corps.Repere);
                                Body2 Corps = null;

                                foreach (var c in Piece.eListeCorps(true))
                                    if (c.Name.EndsWith("-<" + nomCorps + ">"))
                                        Corps = c;

                                Corps.eSelect();
                                mdlFichier.FeatureManager.InsertDeleteBody2(true);

                                mdlFichier.EditRebuild3();
                                mdlFichier.LockAllExternalReferences();
                                fonc.UpdateExternalFileReferences((int)swExternalFileReferencesConfig_e.swExternalFileReferencesCurrentConfig, "", (int)swExternalFileReferencesUpdate_e.swExternalFileReferencesLockAll);
                                mdlFichier.EditRebuild3();

                                mdlFichier.eSauver();
                            }
                            else
                                WindowLog.EcrireF("- {0} erreur", CONSTANTES.PREFIXE_REF_DOSSIER + corps.Repere);

                            App.Sw.CloseDoc(mdlFichier.GetPathName());
                        }
                    }

                    ///////////////////////////////////////////////////////////////////////////////////////////////////////

                    // Petit récap
                    WindowLog.SautDeLigne();
                    WindowLog.EcrireF("Nb de corps unique : {0}", ListeCampagnes[IndiceCampagne].Count);

                    int nbtt = 0;

                    using (var sw = new StreamWriter(FichierNomenclature, false, Encoding.GetEncoding(1252)))
                    {
                        sw.WriteLine(String.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}", "Campagne", "Repere", "Nb", "Type", "Dimension", "Materiau"));

                        foreach (var listecorps in ListeCampagnes.Values)
                        {
                            foreach (var corps in listecorps.Values)
                            {
                                sw.WriteLine(String.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}", corps.Campagne, corps.Repere, corps.Nb, corps.TypeCorps, corps.Dimension, corps.Materiau));

                                if (corps.Campagne == IndiceCampagne)
                                {
                                    nbtt += corps.Nb;
                                    WindowLog.EcrireF("{2} P{0} ×{1}", corps.Repere, corps.Nb, corps.Campagne);
                                }
                            }
                        }

                        
                    }

                    WindowLog.EcrireF("Nb total de corps : {0}", nbtt);

                    MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                }
                catch (Exception e) { this.LogErreur(new Object[] { e }); }
            }

            private Dimension GetParam(ModelDoc2 mdl, Feature fDossier, String nomCfg, ref int indexDimension, out Boolean nouveauDossier)
            {
                Dimension param = null;
                nouveauDossier = false;

                try
                {
                    String nomParam = "";

                    Func<String, String> ExtractNomParam = delegate (String s)
                    {
                        s = s.Replace(CONSTANTES.PREFIXE_REF_DOSSIER + "\"", "").Replace("\"", "");
                        var t = s.Split('@');
                        if (t.Length > 2)
                            return String.Format("{0}@{1}", t[0], t[1]);

                        this.LogErreur(new Object[] { "Pas de parametre dans la reference dossier" });
                        return "";
                    };

                    // On recherche si le dossier contient déjà la propriété RefDossier
                    //      Si non, on ajoute la propriété au dossier selon le modèle suivant :
                    //              P"D1@REPERAGE_DOSSIER@Nom_de_la_piece.SLDPRT"
                    //      Si oui, on récupère le nom du paramètre à configurer

                    CustomPropertyManager PM = fDossier.CustomPropertyManager;
                    String val;
                    if (!PM.ePropExiste(CONSTANTES.REF_DOSSIER))
                    {
                        nouveauDossier = true;
                        nomParam = String.Format("D{0}@{1}", ++indexDimension, CONSTANTES.NOM_ESQUISSE_NUMEROTER);
                        val = String.Format("{0}\"{1}@{2}\"", CONSTANTES.PREFIXE_REF_DOSSIER, nomParam, mdl.eNomAvecExt());
                        var r = PM.ePropAdd(CONSTANTES.REF_DOSSIER, val);
                    }
                    else
                    {
                        String result = ""; Boolean wasResolved, link;
                        var r = PM.Get6(CONSTANTES.REF_DOSSIER, false, out val, out result, out wasResolved, out link);
                        nomParam = ExtractNomParam(val);
                    }

                    PM.ePropAdd(CONSTANTES.DESC_DOSSIER, val);
                    val = String.Format("\"SW-CutListItemName@@@{0}@{1}\"", fDossier.Name, mdl.eNomAvecExt());
                    PM.ePropAdd(CONSTANTES.NOM_DOSSIER, val);

                    param = mdl.Parameter(nomParam);

                    if (nouveauDossier)
                        param.SetSystemValue3(0.5 * 0.001, (int)swSetValueInConfiguration_e.swSetValue_InSpecificConfigurations, nomCfg);
                }
                catch (Exception e) { this.LogErreur(new Object[] { e }); }

                return param;
            }

            private int GetRepere(Dimension param, String nomCfg)
            {
                Double val = (Double)(param.GetSystemValue3((int)swSetValueInConfiguration_e.swSetValue_InSpecificConfigurations, nomCfg)[0]);
                return (int)(val * 1000);
            }

            private void SetRepere(Dimension param, int val, String nomCfg)
            {
                param.SetSystemValue3(val * 0.001, (int)swSetValueInConfiguration_e.swSetValue_InSpecificConfigurations, nomCfg);
            }

            private void SupprimerDefBloc(ModelDoc2 mdl, String cheminbloc)
            {
                var TabDef = (Object[])mdl.SketchManager.GetSketchBlockDefinitions();
                if (TabDef.IsRef())
                {
                    foreach (SketchBlockDefinition blocdef in TabDef)
                    {
                        if (blocdef.FileName == cheminbloc)
                        {
                            Feature d = blocdef.GetFeature();
                            d.eSelect();
                            mdl.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                            mdl.eEffacerSelection();
                            break;
                        }
                    }
                }
            }

            private String CheminBlocEsquisseNumeroter()
            {
                return Sw.CheminBloc(CONSTANTES.NOM_BLOCK_ESQUISSE_NUMEROTER);
            }

            private Feature EsquisseRepere(ModelDoc2 mdl, Boolean creer = true)
            {
                // On recherche l'esquisse contenant les parametres
                Feature Esquisse = mdl.eChercherFonction(fc => { return fc.Name == CONSTANTES.NOM_ESQUISSE_NUMEROTER; });

                if (Esquisse.IsNull() && creer)
                {
                    var SM = mdl.SketchManager;

                    // On recherche le chemin du bloc
                    String cheminbloc = CheminBlocEsquisseNumeroter();

                    if (String.IsNullOrWhiteSpace(cheminbloc))
                        return null;

                    // On supprime la definition du bloc
                    SupprimerDefBloc(mdl, cheminbloc);

                    // On recherche le plan de dessus, le deuxième dans la liste des plans de référence
                    Feature Plan = mdl.eListeFonctions(fc => { return fc.GetTypeName2() == FeatureType.swTnRefPlane; })[1];

                    // Selection du plan et création de l'esquisse
                    Plan.eSelect();
                    SM.InsertSketch(true);
                    SM.AddToDB = false;
                    SM.DisplayWhenAdded = true;

                    mdl.eEffacerSelection();

                    // On récupère la fonction de l'esquisse
                    Esquisse = mdl.Extension.GetLastFeatureAdded();

                    // On insère le bloc
                    MathUtility Mu = App.Sw.GetMathUtility();
                    MathPoint Origine = Mu.CreatePoint(new double[] { 0, 0, 0 });
                    var def = SM.MakeSketchBlockFromFile(Origine, cheminbloc, false, 1, 0);

                    // On récupère la première instance
                    // et on l'explose
                    var Tab = (Object[])def.GetInstances();
                    var ins = (SketchBlockInstance)Tab[0];
                    SM.ExplodeSketchBlockInstance(ins);

                    // Fermeture de l'esquisse
                    SM.AddToDB = false;
                    SM.DisplayWhenAdded = true;
                    SM.InsertSketch(true);

                    //// On supprime la definition du bloc
                    //SupprimerDefBloc(mdl, cheminbloc);

                    // On renomme l'esquisse
                    Esquisse.Name = CONSTANTES.NOM_ESQUISSE_NUMEROTER;

                    mdl.eEffacerSelection();

                    // On l'active dans toutes les configurations
                    Esquisse.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, null);
                }

                if (Esquisse.IsRef())
                {
                    // On selectionne l'esquisse, on la cache
                    // et on la masque dans le FeatureMgr
                    // elle ne sera pas du tout acessible par l'utilisateur
                    Esquisse.eSelect();
                    mdl.BlankSketch();
                    Esquisse.SetUIState((int)swUIStates_e.swIsHiddenInFeatureMgr, true);
                    mdl.eEffacerSelection();

                    mdl.EditRebuild3();
                }

                return Esquisse;
            }

            private void NettoyerModele()
            {
                WindowLog.Ecrire("Nettoyer les modeles");
                List<ModelDoc2> ListeMdl = new List<ModelDoc2>();
                if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                    ListeMdl.Add(MdlBase);
                else
                    ListeMdl = MdlBase.eAssemblyDoc().eListeModeles();

                Predicate<Feature> Test = delegate (Feature f)
                {
                    BodyFolder dossier = f.GetSpecificFeature2();
                    if (dossier.IsRef() && dossier.eNbCorps() > 0)
                        return true;

                    return false;
                };

                foreach (var mdl in ListeMdl)
                {
                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                    {
                        mdl.ePropSuppr(CONST_PRODUCTION.ID_PIECE);
                        mdl.ePropSuppr(CONST_PRODUCTION.MAX_INDEXDIM);

                        // On supprime la definition du bloc
                        SupprimerDefBloc(mdl, CheminBlocEsquisseNumeroter());
                    }

                    foreach (var cfg in mdl.eListeNomConfiguration())
                    {
                        mdl.ShowConfiguration2(cfg);
                        mdl.EditRebuild3();
                        var Piece = mdl.ePartDoc();

                        mdl.ePropSuppr(CONST_PRODUCTION.ID_CONFIG, cfg);
                        mdl.ePropSuppr(CONST_PRODUCTION.ID_DOSSIERS, cfg);

                        foreach (var f in Piece.eListeDesFonctionsDePiecesSoudees(Test))
                        {
                            CustomPropertyManager PM = f.CustomPropertyManager;
                            PM.Delete2(CONSTANTES.REF_DOSSIER);
                            PM.Delete2(CONSTANTES.DESC_DOSSIER);
                            PM.Delete2(CONSTANTES.NOM_DOSSIER);
                        }
                    }

                    if (mdl.GetPathName() != MdlBase.GetPathName())
                        App.Sw.CloseDoc(mdl.GetPathName());
                }

                int errors = 0;
                int warnings = 0;
                MdlBase.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent + (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, ref errors, ref warnings);

                WindowLog.Ecrire("\nNettoyage terminé");
            }
        }
    }
}


