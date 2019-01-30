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
        public Boolean CombinerAvecCampagnePrecedente = false;
        public Boolean CreerDvp = false;
        public eTypeCorps FiltrerCorps = eTypeCorps.Piece;

        public ListeSortedCorps ListeCorps = new ListeSortedCorps();

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
                // Si aucun corps n'a déjà été repéré, on reinitialise tout
                if (ListeCorps.Count == 0)
                {
                    // On supprime tout les fichiers
                    foreach (FileInfo file in new DirectoryInfo(MdlBase.pDossierPiece()).GetFiles())
                        file.Delete();
                }

                if (ReinitCampagneActuelle && (ListeCorps.Count > 0))
                {
                    // On supprime les repères de la campagne actuelle

                    // On recherche les repères appartenant aux campagnes précédentes
                    // pour ne pas supprimer les fichiers
                    // Si la somme des quantités des campagnes précédente est superieure à 0
                    // on garde le repère
                    ListeSortedCorps FichierAsauvegarder = new ListeSortedCorps();
                    FichierAsauvegarder.CampagneDepartDecompte = ListeCorps.CampagneDepartDecompte;

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

                    // On nettoie les fichiers précedement crées
                    foreach (var corps in ListeCorps.Values)
                    {
                        if (FichierAsauvegarder.ContainsKey(corps.Repere)) continue;

                        File.Delete(corps.CheminFichierRepere);
                        File.Delete(corps.CheminFichierApercu);
                    }

                    ListeCorps = FichierAsauvegarder;

                }

                // On supprime les campagnes superieures à l'indice actuelle
                foreach (var corps in ListeCorps.Values)
                {
                    corps.InitCampagne(IndiceCampagne);

                    for (int i = IndiceCampagne; i < corps.Campagne.Keys.Max(); i++)
                    {
                        if (corps.Campagne.ContainsKey(i + 1))
                            corps.Campagne.Remove(i + 1);
                    }
                }

                // On charge les corps existant à partir des fichiers
                if (CombinerCorpsIdentiques && CombinerAvecCampagnePrecedente && (ListeCorps.Count > 0))
                {
                    WindowLog.SautDeLigne();
                    WindowLog.EcrireF("Chargement des corps existants ({0}):", ListeCorps.Count);

                    foreach (var corps in ListeCorps.Values)
                    {
                        if (File.Exists(corps.CheminFichierRepere))
                        {
                            WindowLog.EcrireF("- {0}", corps.RepereComplet);
                            ModelDoc2 mdl = Sw.eOuvrir(corps.CheminFichierRepere);
                            mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                            var Piece = mdl.ePartDoc();
                            ListeCorps[corps.Repere].SwCorps = Piece.ePremierCorps();
                        }
                    }
                }

                ////////////////////////////////// DEBUT DU REPERAGE ////////////////////////////////////////////////////

                MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                // On recherche l'indice de repere max
                if (ListeCorps.Count > 0)
                    _GenRepereDossier = ListeCorps.Keys.Max();

                // On liste les composants
                var ListeComposants = MdlBase.pListerComposants();

                // On boucle sur les modeles
                foreach (var mdl in ListeComposants.Keys)
                {
                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                    // On met à jour les options
                    AppliqueOptionListeDePiecesSoudees(mdl);

                    // On crée l'esquisse pour le reperage des dossiers
                    mdl.pEsquisseRepere();

                    // Si le modele est a repérer complètement
                    Boolean InitModele = true;
                    // On recherche l'index de la dimension maximum
                    int IndexDimension = 0;
                    // On liste les dossiers déja traité pour l'attribution des nouveaux index de dimension
                    HashSet<int> HashPieceIdDossiers = new HashSet<int>();

                    // Les données sont stockées dans des propriétés du modèle
                    // Le nom du modèle est stocké dans une propriété, si le modèle est copié
                    // la propriété n'est plus valable, on force le repérage
                    // On récupère également le dernier indice de la dimension utilisée

                    if (mdl.ePropExiste(CONST_PRODUCTION.ID_PIECE) && (mdl.eGetProp(CONST_PRODUCTION.ID_PIECE) == mdl.eNomSansExt()))
                    {
                        InitModele = false;
                        if (mdl.ePropExiste(CONST_PRODUCTION.MAX_INDEXDIM))
                            IndexDimension = mdl.eGetProp(CONST_PRODUCTION.MAX_INDEXDIM).eToInteger();

                        if (mdl.ePropExiste(CONST_PRODUCTION.PIECE_ID_DOSSIERS))
                        {
                            var tab = mdl.eGetProp(CONST_PRODUCTION.PIECE_ID_DOSSIERS).Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var id in tab)
                                HashPieceIdDossiers.Add(id.eToInteger());
                        }
                    }

                    foreach (var nomCfg in ListeComposants[mdl].Keys)
                    {
                        mdl.ShowConfiguration2(nomCfg);
                        mdl.EditRebuild3();
                        WindowLog.SautDeLigne();
                        WindowLog.EcrireF("{0} \"{1}\"", mdl.eNomSansExt(), nomCfg);

                        HashSet<int> HashConfigIdDossiers = new HashSet<int>();

                        Boolean InitConfig = true;

                        int IdCfg = mdl.GetConfigurationByName(nomCfg).GetID();

                        // Idem modèle, on stock l'id de la config dans une propriété.
                        // Si une nouvelle config est crée, la valeur de cette propriété devient caduc,
                        // on repère alors les dossiers
                        // On en profite pour récupérer la liste des ids de dossiers déjà traité dans les précédentes
                        // campagne de repérage
                        if (!InitModele && mdl.ePropExiste(CONST_PRODUCTION.ID_CONFIG, nomCfg) && (mdl.eGetProp(CONST_PRODUCTION.ID_CONFIG, nomCfg) == IdCfg.ToString()))
                        {
                            InitConfig = false;
                            if (mdl.ePropExiste(CONST_PRODUCTION.CONFIG_ID_DOSSIERS, nomCfg))
                            {
                                var tab = mdl.eGetProp(CONST_PRODUCTION.CONFIG_ID_DOSSIERS, nomCfg).Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                foreach (var id in tab)
                                    HashConfigIdDossiers.Add(id.eToInteger());
                            }
                        }

                        var piece = mdl.ePartDoc();
                        var ListeDossier = piece.eListeDesFonctionsDePiecesSoudees(
                            swD =>
                            {
                                BodyFolder Dossier = swD.GetSpecificFeature2();

                                // Si le dossier est la racine d'un sous-ensemble soudé, il n'y a rien dedans
                                if (Dossier.IsRef() && (Dossier.eNbCorps() > 0) &&
                                FiltrerCorps.HasFlag(Dossier.eTypeDeDossier()))
                                    return true;

                                return false;
                            }
                            );

                        var NbConfig = ListeComposants[mdl][nomCfg];

                        foreach (var fDossier in ListeDossier)
                        {
                            BodyFolder Dossier = fDossier.GetSpecificFeature2();
                            int IdDossier = fDossier.GetID();

                            Dimension param = null;

                            if (!HashPieceIdDossiers.Contains(IdDossier))
                                param = CreerParam(mdl, fDossier, nomCfg, ++IndexDimension);
                            else
                                param = GetParam(mdl, fDossier, nomCfg);

                            var SwCorps = Dossier.ePremierCorps();
                            var NomCorps = SwCorps.Name;
                            var MateriauCorps = SwCorps.eGetMateriauCorpsOuPiece(piece, nomCfg);
                            eTypeCorps TypeCorps = Dossier.eTypeDeDossier();
                            var nbCorps = Dossier.eNbCorps() * NbConfig;

                            int Repere = -1;

                            if (CombinerCorpsIdentiques)
                            {
                                // On recherche s'il existe des corps identiques
                                // Si oui, on applique le même repère au parametre

                                foreach (var CorpsTest in ListeCorps.Values)
                                {
                                    if ((CombinerAvecCampagnePrecedente || CorpsTest.Campagne.ContainsKey(IndiceCampagne)) &&
                                        CorpsTest.SwCorps.IsRef() &&
                                        (MateriauCorps == CorpsTest.Materiau) &&
                                        (TypeCorps == CorpsTest.TypeCorps) &&
                                        SwCorps.eEstSemblable(CorpsTest.SwCorps))
                                    {
                                        Repere = CorpsTest.Repere;
                                        SetRepere(param, CorpsTest.Repere, nomCfg);
                                        break;
                                    }
                                }
                            }

                            // Initialisation du repère
                            if (Repere.EstNegatif())
                            {
                                // Si on est pas en mode "Combiner les corps"
                                // on recupère le repère du dossier
                                // Sinon c'est forcément un nouveau repère
                                if (!CombinerCorpsIdentiques)
                                    Repere = GetRepere(param, nomCfg);

                                // Création d'un nouveau repère suivant conditions
                                // Dans tous les cas, si la clé est négative, on crée un nouveau repère
                                if (Repere.EstNegatif() ||
                                    InitConfig ||
                                    !HashConfigIdDossiers.Contains(IdDossier) ||
                                    !ListeCorps.ContainsKey(Repere))
                                {
                                    Repere = GenRepereDossier;
                                    SetRepere(param, Repere, nomCfg);
                                }
                            }

                            // Initialisation du corps
                            Corps corps = null;
                            if (!ListeCorps.ContainsKey(Repere))
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

                            HashPieceIdDossiers.Add(IdDossier);
                            HashConfigIdDossiers.Add(IdDossier);

                            WindowLog.EcrireF(" - {1} -> {0}", fDossier.Name, corps.RepereComplet);
                        }
                        mdl.ePropAdd(CONST_PRODUCTION.ID_CONFIG, IdCfg, nomCfg);
                        mdl.ePropAdd(CONST_PRODUCTION.CONFIG_ID_DOSSIERS, String.Join(" ", HashConfigIdDossiers), nomCfg);
                    }
                    mdl.ePropAdd(CONST_PRODUCTION.ID_PIECE, mdl.eNomSansExt());
                    mdl.ePropAdd(CONST_PRODUCTION.PIECE_ID_DOSSIERS, String.Join(" ", HashPieceIdDossiers));
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
                foreach (var corps in ListeCorps.Values)
                    App.Sw.CloseDoc(corps.CheminFichierRepere);

                WindowLog.SautDeLigne();
                if (ListeCorps.Count > 0)
                    WindowLog.EcrireF("Nb de repères : {0}", ListeCorps.Keys.Max());
                else
                    WindowLog.Ecrire("Aucun corps repéré");

                // S'il n'y a aucun corps, on se barre
                if (ListeCorps.Count == 0)
                    return;

                ////////////////////////////////// EXPORTER LES CORPS /////////////////////////////////////////////////

                WindowLog.SautDeLigne();
                WindowLog.Ecrire("Export des corps :");

                foreach (var corps in ListeCorps.Values)
                {
                    if (corps.Modele.IsNull()) continue;

                    // Si on est pas en mode "Combiner corps identique" et que le fichier existe
                    // on le supprime pour le mettre à jour, sinon on peut se retrouver
                    // avec des fichiers ne correpondants pas au corps
                    if (!CombinerCorpsIdentiques && File.Exists(corps.CheminFichierRepere))
                        File.Delete(corps.CheminFichierRepere);

                    // Si le fichier existe, on passe au suivant
                    if (File.Exists(corps.CheminFichierRepere))
                        continue;

                    WindowLog.EcrireF("- {0} exporté", CONSTANTES.PREFIXE_REF_DOSSIER + corps.Repere);

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
                    mdlFichier.FixerProp(corps.RepereComplet);

                    if ((corps.TypeCorps == eTypeCorps.Tole) && CreerDvp)
                        corps.pCreerDvp(MdlBase.pDossierPiece(), false);

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

                // Petit récap
                WindowLog.SautDeLigne();
                WindowLog.EcrireF("Nb de corps unique : {0}", ListeCorps.Count);

                int nbtt = 0;

                foreach (var corps in ListeCorps.Values)
                {
                    nbtt += corps.Campagne[IndiceCampagne];
                    if (corps.Campagne[IndiceCampagne] > 0)
                        WindowLog.EcrireF("{2} P{0} ×{1}", corps.Repere, corps.Campagne[IndiceCampagne], IndiceCampagne);
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
            Bitmap bmp = RedimensionnerImage(100, 100, CheminImg);
            bmp.Save(CheminImg);
        }

        public Bitmap RedimensionnerImage(int newWidth, int newHeight, string stPhotoPath)
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

        private Dimension GetParam(ModelDoc2 mdl, Feature fDossier, String nomCfg)
        {
            Dimension param = null;

            try
            {
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

                String result = ""; Boolean wasResolved, link;
                var r = PM.Get6(CONSTANTES.REF_DOSSIER, false, out val, out result, out wasResolved, out link);
                String nomParam = ExtractNomParam(val);

                PM.ePropAdd(CONSTANTES.DESC_DOSSIER, val);
                val = String.Format("\"SW-CutListItemName@@@{0}@{1}\"", fDossier.Name, mdl.eNomAvecExt());
                PM.ePropAdd(CONSTANTES.NOM_DOSSIER, val);

                param = mdl.Parameter(nomParam);
                param.SetSystemValue3(0.5 * 0.001, (int)swSetValueInConfiguration_e.swSetValue_InSpecificConfigurations, nomCfg);
            }
            catch (Exception e) { this.LogErreur(new Object[] { e }); }

            return param;
        }

        private Dimension CreerParam(ModelDoc2 mdl, Feature fDossier, String nomCfg, int indexDimension)
        {
            Dimension param = null;

            try
            {
                // On recherche si le dossier contient déjà la propriété RefDossier
                //      Si non, on ajoute la propriété au dossier selon le modèle suivant :
                //              P"D1@REPERAGE_DOSSIER@Nom_de_la_piece.SLDPRT"
                //      Si oui, on récupère le nom du paramètre à configurer

                CustomPropertyManager PM = fDossier.CustomPropertyManager;
                String val;

                String nomParam = String.Format("D{0}@{1}", indexDimension, CONSTANTES.NOM_ESQUISSE_NUMEROTER);
                val = String.Format("{0}\"{1}@{2}\"", CONSTANTES.PREFIXE_REF_DOSSIER, nomParam, mdl.eNomAvecExt());
                var r = PM.ePropAdd(CONSTANTES.REF_DOSSIER, val);

                PM.ePropAdd(CONSTANTES.DESC_DOSSIER, val);
                val = String.Format("\"SW-CutListItemName@@@{0}@{1}\"", fDossier.Name, mdl.eNomAvecExt());
                PM.ePropAdd(CONSTANTES.NOM_DOSSIER, val);

                param = mdl.Parameter(nomParam);
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
            Double val = (Double)param.GetSystemValue3((int)swSetValueInConfiguration_e.swSetValue_InSpecificConfigurations, nomCfg)[0];
            if (val == 0.5)
                val = -1;
            else
                val *= 1000;

            return (int)val;
        }

        private void SetRepere(Dimension param, int val, String nomCfg)
        {
            param.SetSystemValue3(val * 0.001, (int)swSetValueInConfiguration_e.swSetValue_InSpecificConfigurations, nomCfg);
        }
    }
}


