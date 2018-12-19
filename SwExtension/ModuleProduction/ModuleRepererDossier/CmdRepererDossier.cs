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

            public Boolean ReinitCampagneActuelle = false;
            public Boolean MajCampagnePrecedente = false;
            public Boolean CombinerCorpsIdentiques = false;
            public Boolean CombinerAvecCampagne = false;
            public Boolean ExporterFichierCorps = false;
            public String FichierNomenclature = "";
            public String CheminDossierPieces = "";
            public SortedDictionary<int, Corps> ListeCorpsExistant = new SortedDictionary<int, Corps>();
            public SortedDictionary<int, String> ListeCorpsCharge = new SortedDictionary<int, String>();
            private String ExtPiece = eTypeDoc.Piece.GetEnumInfo<ExtFichier>();

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

                    // Si aucun corps n'a déjà été repéré, on reinitialise tout
                    if (ListeCorpsExistant.Count == 0)
                    {
                        NettoyerModele();

                        // On supprime tout les fichiers
                        foreach (FileInfo file in new DirectoryInfo(CheminDossierPieces).GetFiles())
                            file.Delete();
                    }

                    if (ReinitCampagneActuelle && (ListeCorpsExistant.Count > 0))
                    {
                        // On supprime les repères de la campagne actuelle

                        // On recherche les repères appartenant aux campagnes précédentes
                        // pour ne pas supprimer les fichiers
                        // Si la somme des quantités des campagnes précédente est superieure à 0
                        // on garde le repère
                        SortedDictionary<int, Corps> FichierAsauvegarder = new SortedDictionary<int, Corps>();
                        foreach (var corps in ListeCorpsExistant.Values)
                        {
                            int nb = 0;
                            foreach (var camp in corps.Campagne)
                            {
                                if (camp.Key < IndiceCampagne)
                                    nb += camp.Value;
                                else
                                    break;
                            }

                            if (nb > 0)
                                FichierAsauvegarder.Add(corps.Repere, corps);
                        }

                        // On nettoie les fichiers précedement crées
                        foreach (var repere in ListeCorpsExistant.Keys)
                        {
                            if (FichierAsauvegarder.ContainsKey(repere)) continue;

                            String fichier = Path.Combine(CheminDossierPieces, CONSTANTES.PREFIXE_REF_DOSSIER + repere + ExtPiece);
                            if (File.Exists(fichier))
                                File.Delete(fichier);
                        }

                        ListeCorpsExistant = FichierAsauvegarder;
                    }

                    // On initialise les quantités de cette campagne à 0;
                    // On supprime les campagnes superieures à l'indice actuelle
                    foreach (var corps in ListeCorpsExistant.Values)
                    {
                        corps.InitCampagne(IndiceCampagne);

                        for (int i = IndiceCampagne; i < corps.Campagne.Keys.Max(); i++)
                        {
                            if (corps.Campagne.ContainsKey(i + 1))
                                corps.Campagne.Remove(i + 1);
                        }
                    }

                    // On charge les corps existant à partir des fichiers
                    if (CombinerCorpsIdentiques && CombinerAvecCampagne && (ListeCorpsExistant.Count > 0))
                    {
                        WindowLog.SautDeLigne();
                        WindowLog.EcrireF("Chargement des corps existants ({0}):", ListeCorpsExistant.Count);

                        foreach (FileInfo file in new DirectoryInfo(CheminDossierPieces).GetFiles("*" + ExtPiece))
                        {
                            int rep = Path.GetFileNameWithoutExtension(file.Name).Replace(CONSTANTES.PREFIXE_REF_DOSSIER, "").eToInteger();

                            if (ListeCorpsExistant.ContainsKey(rep))
                            {
                                WindowLog.EcrireF("- {0}", Path.GetFileNameWithoutExtension(file.Name));
                                ListeCorpsCharge.Add(rep, file.FullName);
                                ModelDoc2 mdl = Sw.eOuvrir(file.FullName);
                                mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                                var Piece = mdl.ePartDoc();
                                ListeCorpsExistant[rep].SwCorps = Piece.ePremierCorps();
                            }
                        }
                    }

                    MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                    // On recherche l'indice de repere max
                    if (ListeCorpsExistant.Count > 0)
                        _GenRepereDossier = ListeCorpsExistant.Keys.Max();

                    // On liste les composants
                    var ListeComposants = MdlBase.ListerComposants(false);

                    // On boucle sur les modeles
                    foreach (var mdl in ListeComposants.Keys)
                    {
                        mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                        // On met à jour les options
                        AppliqueOptionListeDePiecesSoudees(mdl);

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

                        foreach (var nomCfg in ListeComposants[mdl].Keys)
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

                            var NbConfig = ListeComposants[mdl][nomCfg];

                            foreach (var fDossier in ListeDossier)
                            {
                                BodyFolder Dossier = fDossier.GetSpecificFeature2();
                                int IdDossier = fDossier.GetID();

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

                                    foreach (var CorpsTest in ListeCorpsExistant.Values)
                                    {
                                        if ((CombinerAvecCampagne || CorpsTest.Campagne.ContainsKey(IndiceCampagne)) &&
                                            CorpsTest.SwCorps.IsRef() &&
                                            (MateriauCorps == CorpsTest.Materiau) &&
                                            (TypeCorps == CorpsTest.TypeCorps) &&
                                            SwCorps.eEstSemblable(CorpsTest.SwCorps))
                                        {
                                            Repere = CorpsTest.Repere;
                                            SetRepere(param, CorpsTest.Repere, nomCfg);
                                            Combiner = true;
                                            break;
                                        }
                                    }
                                }

                                
                                if (!Combiner)
                                {
                                    Repere = GetRepere(param, nomCfg);

                                    // Création d'un nouveau repère si
                                    if (InitConfig ||
                                        NouveauDossier ||
                                        !ListIdDossiers.Contains(IdDossier) ||
                                        !ListeCorpsExistant.ContainsKey(Repere))
                                    {
                                        Repere = GenRepereDossier;
                                        SetRepere(param, Repere, nomCfg);
                                    }
                                }

                                Corps corps = null;
                                if (NouveauDossier || !ListeCorpsExistant.ContainsKey(Repere))
                                {
                                    corps = new Corps(SwCorps, TypeCorps, MateriauCorps);
                                    corps.InitCampagne(IndiceCampagne);
                                    ListeCorpsExistant.Add(Repere, corps);
                                }
                                else
                                    corps = ListeCorpsExistant[Repere];

                                corps.Campagne[IndiceCampagne] += nbCorps;
                                corps.Repere = Repere;
                                corps.InitDimension(Dossier, SwCorps);
                                corps.AjouterModele(mdl, nomCfg, IdDossier, NomCorps);

                                ListIdDossiers.Add(IdDossier);

                                WindowLog.EcrireF("     {0} -> {1}", fDossier.Name, corps.Repere);
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

                        foreach (var corps in ListeCorpsExistant.Values)
                        {
                            if (corps.Modele.IsNull()) continue;

                            var cheminFichier = Path.Combine(CheminDossierPieces, CONSTANTES.PREFIXE_REF_DOSSIER + corps.Repere + ExtPiece);
                            if (File.Exists(cheminFichier)) continue;

                            corps.Modele.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                            corps.Modele.ShowConfiguration2(corps.NomConfig);
                            corps.Modele.EditRebuild3();

                            var mdlFichier = Sw.eCreerDocument(CheminDossierPieces, CONSTANTES.PREFIXE_REF_DOSSIER + corps.Repere, eTypeDoc.Piece);
                            mdlFichier.eActiver();
                            var Piece = mdlFichier.ePartDoc();
                            var fonc = Piece.InsertPart3(corps.Modele.GetPathName(), 90241, corps.NomConfig);
                            if (fonc.IsRef())
                            {
                                WindowLog.EcrireF("- {0} exporté", CONSTANTES.PREFIXE_REF_DOSSIER + corps.Repere);
                                Body2 Corps = null;

                                foreach (var c in Piece.eListeCorps(true))
                                    if (c.Name.EndsWith("-<" + corps.NomCorps + ">"))
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
                    WindowLog.EcrireF("Nb de corps unique : {0}", ListeCorpsExistant.Count);

                    int nbtt = 0;

                    using (var sw = new StreamWriter(FichierNomenclature, false, Encoding.GetEncoding(1252)))
                    {
                        sw.WriteLine(Corps.Entete(IndiceCampagne));

                        foreach (var corps in ListeCorpsExistant.Values)
                        {
                            sw.WriteLine(corps.ToString());
                            nbtt += corps.Campagne[IndiceCampagne];
                            WindowLog.EcrireF("{2} P{0} ×{1}", corps.Repere, corps.Campagne[IndiceCampagne], IndiceCampagne);
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

            private void AppliqueOptionListeDePiecesSoudees(ModelDoc2 mdl)
            {
                mdl.Extension.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swDisableDerivedConfigurations)), 0, false);
                mdl.Extension.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swWeldmentRenameCutlistDescriptionPropertyValue)), 0, true);
                mdl.Extension.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swWeldmentCollectIdenticalBodies)), 0, true);
                mdl.Extension.SetUserPreferenceString(((int)(swUserPreferenceStringValue_e.swSheetMetalDescription)), 0, "Tôle");
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
                    if (mdl.TypeDoc() != eTypeDoc.Piece) continue;

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


