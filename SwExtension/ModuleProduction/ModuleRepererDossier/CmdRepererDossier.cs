using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;

namespace ModuleProduction.ModuleRepererDossier
{
    public class CmdRepererDossier : Cmd
    {
        public ModelDoc2 MdlBase = null;
        public int IndiceCampagne = 0;

        public Boolean ReinitCampagneActuelle = false;
        public Boolean CombinerCorpsIdentiques = false;
        public Boolean CombinerAvecCampagne = false;
        public Boolean CreerDvp = false;

        public ListeSortedCorps ListeCorps = new ListeSortedCorps();
        public SortedDictionary<int, String> ListeCorpsCharge = new SortedDictionary<int, String>();

        private int _GenRepereDossier = 0;
        private int GenRepereDossier { get { return ++_GenRepereDossier; } }

        /// <summary>
        /// Pour pouvoir obtenir une r�f�rence unique pour chaque dossier de corps, identiques
        /// dans l'assemblage, on passe par la cr�ation d'une propri�t� dans chaque dossier.
        /// Cette propri�t� est li�e � une cote dans une esquisse dont la valeur est �gale au rep�re.
        /// Suivant la configuration, la valeur de la cote peut changer et donc le rep�re du dossier.
        /// C'est le seul moyen pour avoir un lien entre les dossiers et la configuration du mod�le.
        /// Les propri�t�s des dossiers ne sont pas configurables.
        /// </summary>
        protected override void Command()
        {
            try
            {
                // Si aucun corps n'a d�j� �t� rep�r�, on reinitialise tout
                if (ListeCorps.Count == 0)
                {
                    // On supprime tout les fichiers
                    foreach (FileInfo file in new DirectoryInfo(MdlBase.pDossierPiece()).GetFiles())
                        file.Delete();
                }

                if (ReinitCampagneActuelle && (ListeCorps.Count > 0))
                {
                    // On supprime les rep�res de la campagne actuelle

                    // On recherche les rep�res appartenant aux campagnes pr�c�dentes
                    // pour ne pas supprimer les fichiers
                    // Si la somme des quantit�s des campagnes pr�c�dente est superieure � 0
                    // on garde le rep�re
                    ListeSortedCorps FichierAsauvegarder = new ListeSortedCorps();
                    foreach (var corps in ListeCorps.Values)
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

                    // On nettoie les fichiers pr�cedement cr�es
                    foreach (var corps in ListeCorps.Values)
                    {
                        if (FichierAsauvegarder.ContainsKey(corps.Repere)) continue;

                        File.Delete(corps.CheminFichierRepere);
                        File.Delete(corps.CheminFichierApercu);
                    }

                    ListeCorps = FichierAsauvegarder;
                }

                // On supprime les campagnes superieures � l'indice actuelle
                foreach (var corps in ListeCorps.Values)
                {
                    corps.InitCampagne(IndiceCampagne);

                    for (int i = IndiceCampagne; i < corps.Campagne.Keys.Max(); i++)
                    {
                        if (corps.Campagne.ContainsKey(i + 1))
                            corps.Campagne.Remove(i + 1);
                    }
                }

                // On charge les corps existant � partir des fichiers
                if (CombinerCorpsIdentiques && CombinerAvecCampagne && (ListeCorps.Count > 0))
                {
                    WindowLog.SautDeLigne();
                    WindowLog.EcrireF("Chargement des corps existants ({0}):", ListeCorps.Count);

                    foreach (FileInfo file in new DirectoryInfo(MdlBase.pDossierPiece()).GetFiles("*" + OutilsProd.ExtPiece))
                    {
                        int rep = Path.GetFileNameWithoutExtension(file.Name).Replace(CONSTANTES.PREFIXE_REF_DOSSIER, "").eToInteger();

                        if (ListeCorps.ContainsKey(rep))
                        {
                            WindowLog.EcrireF("- {0}", Path.GetFileNameWithoutExtension(file.Name));
                            ListeCorpsCharge.Add(rep, file.FullName);
                            ModelDoc2 mdl = Sw.eOuvrir(file.FullName);
                            mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                            var Piece = mdl.ePartDoc();
                            ListeCorps[rep].SwCorps = Piece.ePremierCorps();
                        }
                    }
                }

                MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                // On recherche l'indice de repere max
                if (ListeCorps.Count > 0)
                    _GenRepereDossier = ListeCorps.Keys.Max();

                // On liste les composants
                var ListeComposants = MdlBase.pListerComposants(false);

                // On boucle sur les modeles
                foreach (var mdl in ListeComposants.Keys)
                {
                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                    // On met � jour les options
                    AppliqueOptionListeDePiecesSoudees(mdl);

                    // On cr�e l'esquisse pour le reperage des dossiers
                    EsquisseRepere(mdl);

                    // Si le modele est a rep�rer compl�tement
                    Boolean InitModele = true;
                    // On recherche l'index de la dimension maximum
                    int IndexDimension = 0;

                    // Les donn�es sont stock�es dans des propri�t�s du mod�le
                    // Le nom du mod�le est stock� dans une propri�t�, si le mod�le est copi�
                    // la propri�t� n'est plus valable, on force le rep�rage
                    // On r�cup�re �galement le dernier indice de la dimension utilis�e
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

                        // Idem mod�le, on stock l'id de la config dans une propri�t�.
                        // Si une nouvelle config est cr�e, la valeur de cette propri�t� devient caduc,
                        // on rep�re alors les dossiers
                        // On en profite pour r�cup�rer la liste des ids de dossiers d�j� trait� dans les pr�c�dentes
                        // campagne de rep�rage
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

                                // Si le dossier est la racine d'un sous-ensemble soud�, il n'y a rien dedans
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
                                // Si oui, on applique le m�me rep�re au parametre

                                foreach (var CorpsTest in ListeCorps.Values)
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

                                // Cr�ation d'un nouveau rep�re si
                                if (InitConfig ||
                                    NouveauDossier ||
                                    !ListIdDossiers.Contains(IdDossier) ||
                                    !ListeCorps.ContainsKey(Repere))
                                {
                                    Repere = GenRepereDossier;
                                    SetRepere(param, Repere, nomCfg);
                                }
                            }

                            Corps corps = null;
                            if (NouveauDossier || !ListeCorps.ContainsKey(Repere))
                            {
                                corps = new Corps(SwCorps, TypeCorps, MateriauCorps, MdlBase);
                                corps.InitCampagne(IndiceCampagne);
                                ListeCorps.Add(Repere, corps);
                            }
                            else
                                corps = ListeCorps[Repere];

                            corps.Campagne[IndiceCampagne] += nbCorps;
                            corps.Repere = Repere;
                            corps.InitCaracteristiques(Dossier, SwCorps);
                            corps.AjouterModele(mdl, nomCfg, IdDossier, NomCorps);

                            ListIdDossiers.Add(IdDossier);

                            WindowLog.EcrireF(" - {1} -> {0}", fDossier.Name, corps.RepereComplet);
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

                // On fermer les fichiers charg�
                foreach (var f in ListeCorpsCharge.Values)
                    App.Sw.CloseDoc(f);

                WindowLog.SautDeLigne();
                WindowLog.EcrireF("Nb de rep�res : {0}", ListeCorps.Keys.Max());

                ////////////////////////////////// EXPORTER LES CORPS /////////////////////////////////////////////////

                WindowLog.SautDeLigne();
                WindowLog.Ecrire("Export des corps :");

                foreach (var corps in ListeCorps.Values)
                {
                    if (corps.Modele.IsNull() || File.Exists(corps.CheminFichierRepere)) continue;

                    WindowLog.EcrireF("- {0} export�", CONSTANTES.PREFIXE_REF_DOSSIER + corps.Repere);

                    corps.Modele.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                    corps.Modele.ShowConfiguration2(corps.NomConfig);
                    corps.Modele.EditRebuild3();

                    // Sauvegarde du fichier de base
                    int Errors = 0, Warning = 0;
                    corps.Modele.Extension.SaveAs(corps.CheminFichierRepere, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)(swSaveAsOptions_e.swSaveAsOptions_Copy | swSaveAsOptions_e.swSaveAsOptions_Silent), null, ref Errors, ref Warning);

                    var mdlFichier = Sw.eOuvrir(corps.CheminFichierRepere, corps.NomConfig);
                    mdlFichier.eActiver();

                    mdlFichier.Extension.BreakAllExternalFileReferences2(true);

                    foreach (var nomCfg in mdlFichier.eListeNomConfiguration())
                        if (nomCfg != corps.NomConfig)
                            mdlFichier.DeleteConfiguration2(nomCfg);

                    var Piece = mdlFichier.ePartDoc();

                    Body2 swCorps = null;

                    foreach (var c in Piece.eListeCorps(false))
                        if (c.Name == corps.NomCorps)
                            swCorps = c;

                    swCorps.eVisible(true);
                    swCorps.eSelect();
                    mdlFichier.FeatureManager.InsertDeleteBody2(true);

                    Piece.ePremierCorps(false).eVisible(true);
                    mdlFichier.EditRebuild3();
                    mdlFichier.pMasquerEsquisses();

                    if ((corps.TypeCorps == eTypeCorps.Tole) && CreerDvp)
                        corps.CreerDvp(MdlBase.pDossierPiece(), false);

                    mdlFichier.FeatureManager.EditFreeze2((int)swMoveFreezeBarTo_e.swMoveFreezeBarToEnd, "", true, true);

                    if (corps.TypeCorps == eTypeCorps.Tole)
                        OrienterVueTole(mdlFichier);
                    else if (corps.TypeCorps == eTypeCorps.Barre)
                        OrienterVueBarre(mdlFichier);

                    SauverVue(mdlFichier, mdlFichier.GetPathName());
                    mdlFichier.EditRebuild3();
                    mdlFichier.eSauver();

                    App.Sw.CloseDoc(mdlFichier.GetPathName());
                }

                ////////////////////////////////// RECAP /////////////////////////////////////////////////

                // Petit r�cap
                WindowLog.SautDeLigne();
                WindowLog.EcrireF("Nb de corps unique : {0}", ListeCorps.Count);

                int nbtt = 0;

                foreach (var corps in ListeCorps.Values)
                {
                    nbtt += corps.Campagne[IndiceCampagne];
                    if (corps.Campagne[IndiceCampagne] > 0)
                        WindowLog.EcrireF("{2} P{0} �{1}", corps.Repere, corps.Campagne[IndiceCampagne], IndiceCampagne);
                }

                WindowLog.EcrireF("Nb total de corps : {0}", nbtt);

                ////////////////////////////////// SAUVEGARDE DE LA NOMENCLATURE /////////////////////////////////////////////

                ListeCorps.EcrireNomenclature(MdlBase.pDossierPiece(), IndiceCampagne);

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////

                MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                var aff = new AffichageElementWPF(ListeCorps, IndiceCampagne);
                aff.ShowDialog();
            }
            catch (Exception e) { this.LogErreur(new Object[] { e }); }
        }

        private void OrienterVueTole(ModelDoc2 mdl)
        {
            var ListeDepliee = mdl.ePartDoc().eListeFonctionsDepliee();
            if (ListeDepliee.Count > 0)
            {
                var fDepliee = ListeDepliee[0];
                FlatPatternFeatureData fDeplieeInfo = fDepliee.GetDefinition();
                Face2 face = fDeplieeInfo.FixedFace2;
                Surface surface = face.GetSurface();

                Boolean Reverse = face.FaceInSurfaceSense();

                Double[] Param = surface.PlaneParams;

                if (Reverse)
                {
                    Param[0] = Param[0] * -1;
                    Param[1] = Param[1] * -1;
                    Param[2] = Param[2] * -1;
                }

                gVecteur Normale = new gVecteur(Param[0], Param[1], Param[2]);
                MathTransform mtNormale = MathRepere(Normale.MathVector());
                MathTransform mtAxeZ = MathRepere(new gVecteur(1, 1, 1).MathVector()); ;

                MathTransform mtRotate = mtAxeZ.Multiply(mtNormale.Inverse());

                ModelView mv = mdl.ActiveView;
                mv.Orientation3 = mtRotate;
                mv.Activate();
            }
            mdl.ViewZoomtofit2();
            mdl.GraphicsRedraw2();
        }

        private void OrienterVueBarre(ModelDoc2 mdl)
        {
            var corps = mdl.ePartDoc().ePremierCorps();

            var analyse = new AnalyseGeomBarre(corps, mdl);

            MathTransform mtNormale = MathRepere(analyse.PlanSection.Normale.MathVector());
            MathTransform mtAxeZ = MathRepere(new gVecteur(1, 1, 1).MathVector()); ;

            MathTransform mtRotate = mtAxeZ.Multiply(mtNormale.Inverse());

            ModelView mv = mdl.ActiveView;
            mv.Orientation3 = mtRotate;
            mv.Activate();

            mdl.ViewZoomtofit2();
            mdl.GraphicsRedraw2();
        }

        private MathTransform MathRepere(MathVector X)
        {
            MathUtility Mu = App.Sw.GetMathUtility();
            MathVector NormAxeX = null, NormAxeY = null, NormAxeZ = null;

            if (X.ArrayData[0] == 0 && X.ArrayData[2] == 0)
            {
                NormAxeZ = Mu.CreateVector(new Double[] { 0, 1, 0 });
                NormAxeX = Mu.CreateVector(new Double[] { 1, 0, 0 });
                NormAxeY = Mu.CreateVector(new Double[] { 0, 0, -1 });
            }
            else
            {
                NormAxeZ = X.Normalise();
                NormAxeX = Mu.CreateVector(new Double[] { X.ArrayData[2], 0, -1 * X.ArrayData[0] }).Normalise();
                NormAxeY = NormAxeZ.Cross(NormAxeX).Normalise();
            }

            MathVector NormTrans = Mu.CreateVector(new Double[] { 0, 0, 0 });
            MathTransform Mt = Mu.ComposeTransform(NormAxeX, NormAxeY, NormAxeZ, NormTrans, 1);
            return Mt;
        }

        private void SauverVue(ModelDoc2 mdl, String cheminFichier)
        {
            String Dossier = Path.Combine(Path.GetDirectoryName(cheminFichier), CONST_PRODUCTION.DOSSIER_PIECES_APERCU);
            Directory.CreateDirectory(Dossier);
            String CheminImg = Path.Combine(Dossier, Path.GetFileNameWithoutExtension(cheminFichier) + ".bmp");
            mdl.SaveBMP(CheminImg, 0, 0);
            Bitmap bmp = resizeImage(100, 100, CheminImg);
            bmp.Save(CheminImg);
        }

        public Bitmap resizeImage(int newWidth, int newHeight, string stPhotoPath)
        {
            Bitmap img = new Bitmap(stPhotoPath);
            Bitmap imageSource = img.Clone(new Rectangle(0, 0, img.Width, img.Height), PixelFormat.Format32bppRgb);
            Image imgPhoto = imageSource.AutoCrop();
            img.Dispose();
            imageSource.Dispose();

            int sourceWidth = imgPhoto.Width;
            int sourceHeight = imgPhoto.Height;

            //Consider vertical pics
            if (sourceWidth < sourceHeight)
            {
                int buff = newWidth;

                newWidth = newHeight;
                newHeight = buff;
            }

            int sourceX = 0, sourceY = 0, destX = 0, destY = 0;
            float nPercent = 0, nPercentW = 0, nPercentH = 0;

            nPercentW = ((float)newWidth / (float)sourceWidth);
            nPercentH = ((float)newHeight / (float)sourceHeight);
            if (nPercentH < nPercentW)
            {
                nPercent = nPercentH;
                destX = System.Convert.ToInt16((newWidth -
                          (sourceWidth * nPercent)) / 2);
            }
            else
            {
                nPercent = nPercentW;
                destY = System.Convert.ToInt16((newHeight -
                          (sourceHeight * nPercent)) / 2);
            }

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);


            Bitmap bmPhoto = new Bitmap(newWidth, newHeight,
                          PixelFormat.Format32bppRgb);

            bmPhoto.SetResolution(imgPhoto.HorizontalResolution,
                         imgPhoto.VerticalResolution);

            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.Clear(Color.White);
            grPhoto.InterpolationMode = InterpolationMode.HighQualityBicubic;

            grPhoto.DrawImage(imgPhoto,
                new Rectangle(destX, destY, destWidth, destHeight),
                new Rectangle(sourceX, sourceY, sourceWidth, sourceHeight),
                GraphicsUnit.Pixel);

            grPhoto.Dispose();
            imgPhoto.Dispose();
            return bmPhoto;
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

                // On recherche si le dossier contient d�j� la propri�t� RefDossier
                //      Si non, on ajoute la propri�t� au dossier selon le mod�le suivant :
                //              P"D1@REPERAGE_DOSSIER@Nom_de_la_piece.SLDPRT"
                //      Si oui, on r�cup�re le nom du param�tre � configurer

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
            mdl.Extension.SetUserPreferenceString(((int)(swUserPreferenceStringValue_e.swSheetMetalDescription)), 0, "T�le");
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

                // On recherche le plan de dessus, le deuxi�me dans la liste des plans de r�f�rence
                Feature Plan = mdl.eListeFonctions(fc => { return fc.GetTypeName2() == FeatureType.swTnRefPlane; })[1];

                // Selection du plan et cr�ation de l'esquisse
                Plan.eSelect();
                SM.InsertSketch(true);
                SM.AddToDB = false;
                SM.DisplayWhenAdded = true;

                mdl.eEffacerSelection();

                // On r�cup�re la fonction de l'esquisse
                Esquisse = mdl.Extension.GetLastFeatureAdded();

                // On ins�re le bloc
                MathUtility Mu = App.Sw.GetMathUtility();
                MathPoint Origine = Mu.CreatePoint(new double[] { 0, 0, 0 });
                var def = SM.MakeSketchBlockFromFile(Origine, cheminbloc, false, 1, 0);

                // On r�cup�re la premi�re instance
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
    }
}


