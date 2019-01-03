using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Media.Imaging;

namespace ModuleProduction
{
    public static class CONST_PRODUCTION
    {
        public const String DOSSIER_PIECES = "Pieces";
        public const String DOSSIER_PIECES_APERCU = "Apercu";
        public const String FICHIER_NOMENC = "Nomenclature.txt";
        public const String DOSSIER_LASERTOLE = "Laser tole";
        public const String DOSSIER_LASERTUBE = "Laser tube";
        public const String CAMPAGNE_DEPART_DECOMPTE = "CampagneDepartDecompte";
        public const String MAX_INDEXDIM = "MAX_INDEXDIM";
        public const String ID_PIECE = "ID_PIECE";
        public const String ID_CONFIG = "ID_CONFIG";
        public const String ID_DOSSIERS = "ID_DOSSIERS";
    }

    public static class OutilsProd
    {
        public static String Quantite(this ModelDoc2 mdl)
        {
            CustomPropertyManager PM = mdl.Extension.get_CustomPropertyManager("");

            if (mdl.ePropExiste(CONSTANTES.PROPRIETE_QUANTITE))
                return Math.Max(mdl.eProp(CONSTANTES.PROPRIETE_QUANTITE).eToInteger(), 1).ToString();

            return "1";
        }

        public static Boolean pCreerConfigDepliee(this ModelDoc2 mdl, String NomConfigDepliee, String NomConfigPliee)
        {
            var cfg = mdl.ConfigurationManager.AddConfiguration(NomConfigDepliee, NomConfigDepliee, "", 0, NomConfigPliee, "");
            if (cfg.IsRef())
                return true;

            return false;
        }

        public static void pDeplierTole(this PartDoc piece, String nomConfigDepliee)
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

        public static void pPlierTole(this PartDoc piece, String nomConfigPliee)
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

        public static void pMasquerEsquisses(this ModelDoc2 mdl)
        {
            mdl.eParcourirFonctions(
                                    f =>
                                    {
                                        if (f.GetTypeName2() == FeatureType.swTnFlatPattern)
                                            return true;
                                        else if (f.GetTypeName2() == FeatureType.swTnProfileFeature)
                                        {
                                            f.eSelect(false);
                                            mdl.BlankSketch();
                                            mdl.eEffacerSelection();
                                        }
                                        return false;
                                    },
                                    true
                                    );
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
        public static SortedDictionary<ModelDoc2, SortedDictionary<String, int>> pListerComposants(this ModelDoc2 mdlBase, Boolean composantsExterne)
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

        private const String ChaineIndice = "ZYXWVUTSRQPONMLKJIHGFEDCBA";

        public static String pChercherIndice(List<String> liste)
        {
            for (int i = 0; i < ChaineIndice.Length; i++)
            {
                if (liste.Any(d => { return d.EndsWith(" Ind " + ChaineIndice[i]) ? true : false; }))
                    return "Ind " + ChaineIndice[Math.Max(0, i - 1)];
            }

            return "Ind " + ChaineIndice.Last();
        }

        public static String pCreerDossier(this ModelDoc2 mdl, String dossier)
        {
            var chemin = Path.Combine(mdl.eDossier(), dossier);
            if (!Directory.Exists(chemin))
                Directory.CreateDirectory(chemin);

            return chemin;
        }

        public static String pCreerFichierTexte(this ModelDoc2 mdl, String dossier, String fichier)
        {
            var chemin = Path.Combine(mdl.eDossier(), dossier, fichier);
            if (!File.Exists(chemin))
                File.WriteAllText(chemin, "", Encoding.GetEncoding(1252));

            return chemin;
        }

        public static int pRechercherIndiceDossier(this ModelDoc2 mdl, String dossier)
        {
            int indice = 0;
            String chemin = Path.Combine(mdl.eDossier(), dossier);

            if (Directory.Exists(chemin))
                foreach (var d in Directory.EnumerateDirectories(chemin, "*", SearchOption.TopDirectoryOnly))
                    indice = Math.Max(indice, new DirectoryInfo(d).Name.eToInteger());

            return indice;
        }

        public static String pDossierPiece(this ModelDoc2 mdl)
        {
            return Path.Combine(mdl.eDossier(), CONST_PRODUCTION.DOSSIER_PIECES);
        }

        public static String pFichierNomenclature(this ModelDoc2 mdl)
        {
            return Path.Combine(mdl.eDossier(), CONST_PRODUCTION.DOSSIER_PIECES, CONST_PRODUCTION.FICHIER_NOMENC);
        }

        public static String pDossierLaserTole(this ModelDoc2 mdl)
        {
            return Path.Combine(mdl.eDossier(), CONST_PRODUCTION.DOSSIER_LASERTOLE);
        }

        public static String pDossierLaserTube(this ModelDoc2 mdl)
        {
            return Path.Combine(mdl.eDossier(), CONST_PRODUCTION.DOSSIER_LASERTUBE);
        }

        public static ListeSortedCorps pChargerNomenclature(this ModelDoc2 mdl, eTypeCorps type = eTypeCorps.Tous)
        {
            var Liste = new ListeSortedCorps();

            var chemin = mdl.pFichierNomenclature();

            if (File.Exists(chemin))
            {
                using (var sr = new StreamReader(chemin, Encoding.GetEncoding(1252)))
                {
                    // On lit la première ligne contenant l'entête des colonnes
                    String ligne = sr.ReadLine();

                    if (ligne.IsRef())
                    {
                        // On récupère la campagne de départ
                        if (ligne.StartsWith(CONST_PRODUCTION.CAMPAGNE_DEPART_DECOMPTE))
                        {
                            var tab = ligne.Split(new char[] { '\t' });
                            Liste.CampagneDepartDecompte = tab[1].eToInteger();
                            ligne = sr.ReadLine();
                        }

                        while ((ligne = sr.ReadLine()) != null)
                        {
                            if (!String.IsNullOrWhiteSpace(ligne))
                            {
                                var c = new Corps(ligne, mdl);
                                if (type.HasFlag(c.TypeCorps))
                                    Liste.Add(c.Repere, c);
                            }
                        }
                    }
                }
            }

            return Liste;
        }

        public static int pIndiceMaxNomenclature(this ModelDoc2 mdl)
        {
            int index = 0;
            var chemin = mdl.pFichierNomenclature();

            if (File.Exists(chemin))
            {
                using (var sr = new StreamReader(chemin, Encoding.GetEncoding(1252)))
                {
                    // On lit la première ligne contenant l'entête des colonnes
                    String ligne = sr.ReadLine();

                    if (ligne.IsRef())
                    {
                        // On récupère la campagne de départ
                        if (ligne.StartsWith(CONST_PRODUCTION.CAMPAGNE_DEPART_DECOMPTE))
                            ligne = sr.ReadLine();

                        var tab = ligne.Split(new char[] { '\t' });
                        index = tab.Last().eToInteger();
                    }
                }
            }

            return Math.Max(1, index);
        }

        public static ListeSortedCorps pChargerProduction(this ModelDoc2 mdl, String dossierProduction, Boolean mettreAjourCampagne, int campagneDepart = 1)
        {
            var Liste = new ListeSortedCorps();

            Liste.CampagneDepartDecompte = campagneDepart;

            List<String> ListeChemin = new List<String>();

            if (Directory.Exists(dossierProduction))
            {
                var IndiceMax = mdl.pIndiceMaxNomenclature();

                foreach (var d in Directory.EnumerateDirectories(dossierProduction, "*", SearchOption.TopDirectoryOnly))
                {
                    if (!mettreAjourCampagne && ((new DirectoryInfo(d)).Name.eToInteger() == IndiceMax)) continue;

                    var cheminFichier = Path.Combine(d, CONST_PRODUCTION.FICHIER_NOMENC);

                    if (File.Exists(cheminFichier))
                    {
                        using (var sr = new StreamReader(cheminFichier, Encoding.GetEncoding(1252)))
                        {
                            // On lit la première ligne contenant l'entête des colonnes
                            String ligne = sr.ReadLine();

                            if (ligne.IsRef())
                            {
                                // On la split pour récupérer l'indice de la campagne
                                var tab = ligne.Split(new char[] { '\t' });
                                var IndiceCampagne = tab.Last().eToInteger();

                                while ((ligne = sr.ReadLine()) != null)
                                {
                                    if (!String.IsNullOrWhiteSpace(ligne))
                                    {
                                        var c = new Corps(ligne, mdl, IndiceCampagne);

                                        if (Liste.ContainsKey(c.Repere))
                                        {
                                            var tmp = c.Campagne.First();
                                            Liste[c.Repere].Campagne.Add(tmp.Key, tmp.Value);
                                            if (tmp.Key >= campagneDepart)
                                                Liste[c.Repere].Qte += tmp.Value;
                                        }
                                        else
                                        {
                                            Liste.Add(c.Repere, c);

                                            var tmp = c.Campagne.First();
                                            if (tmp.Key >= campagneDepart)
                                                c.Qte = c.Campagne.First().Value;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return Liste;
        }

        public static String ExtPiece = eTypeDoc.Piece.GetEnumInfo<ExtFichier>();

        public static void pCalculerQuantite(this ModelDoc2 mdlBase, ref ListeSortedCorps listeCorps, eTypeCorps typeCorps, List<String> listeMateriaux, List<String> listeDimensions, int indiceCampagne, Boolean mettreAjourCampagne)
        {
            ListeSortedCorps ListeExistant = new ListeSortedCorps();

            if (typeCorps == eTypeCorps.Tole)
                ListeExistant = mdlBase.pChargerProduction(mdlBase.pDossierLaserTole(), mettreAjourCampagne, listeCorps.CampagneDepartDecompte);
            else if (typeCorps == eTypeCorps.Barre)
                ListeExistant = mdlBase.pChargerProduction(mdlBase.pDossierLaserTube(), mettreAjourCampagne, listeCorps.CampagneDepartDecompte);
            else
                return;

            ListeSortedCorps ListeCorpsFiltre = new ListeSortedCorps();
            ListeCorpsFiltre.CampagneDepartDecompte = listeCorps.CampagneDepartDecompte;

            foreach (var corps in listeCorps.Values)
            {
                if ((corps.TypeCorps == typeCorps) &&
                        listeMateriaux.Contains(corps.Materiau) &&
                        listeDimensions.Contains(corps.Dimension)
                        )
                {
                    var qte = corps.Campagne[indiceCampagne];

                    if (ListeExistant.ContainsKey(corps.Repere))
                    {
                        qte = Math.Max(0, qte - ListeExistant[corps.Repere].Qte);

                        // Si la quantité est supérieur à 0
                        // on récupère la différence entre la quantité totale actuelle et
                        // la quantité totale de la précédente campagne
                        if (mettreAjourCampagne && (qte > 0))
                        {
                            var qteCampagnePrecedente = 0;
                            var corpsExistant = ListeExistant[corps.Repere];
                            foreach (var c in corpsExistant.Campagne)
                            {
                                if ((c.Key >= listeCorps.CampagneDepartDecompte) && (c.Key != indiceCampagne))
                                    qteCampagnePrecedente += c.Value;
                            }

                            qte = corps.Campagne[indiceCampagne] - qteCampagnePrecedente;
                        }
                    }

                    corps.Qte = qte;
                    ListeCorpsFiltre.Add(corps.Repere, corps);
                }
            }

            listeCorps = ListeCorpsFiltre;
        }
    }

    [ClassInterface(ClassInterfaceType.None)]
    public class ListeSortedCorps : SortedDictionary<int, Corps>
    {
        private int _CampagneDepartDecompte = 1;
        public int CampagneDepartDecompte { get { return _CampagneDepartDecompte; } set { _CampagneDepartDecompte = value; } }
    }

    public class AnalyseGeomBarre
    {
        public Body2 Corps = null;

        public ModelDoc2 Mdl = null;

        public gPlan PlanSection;
        public gPoint ExtremPoint1;
        public gPoint ExtremPoint2;
        public ListFaceGeom FaceSectionExt = null;
        public List<ListFaceGeom> ListeFaceSectionInt = null;

        public AnalyseGeomBarre(Body2 corps, ModelDoc2 mdl)
        {
            Corps = corps;
            Mdl = mdl;

            AnalyserFaces();
        }

        #region ANALYSE DE LA GEOMETRIE ET RECHERCHE DU PROFIL

        private void AnalyserFaces()
        {
            try
            {
                List<FaceGeom> ListeFaceCorps = new List<FaceGeom>();

                // Tri des faces pour retrouver celles issues de la même
                foreach (var Face in Corps.eListeDesFaces())
                {
                    var faceExt = new FaceGeom(Face);

                    Boolean Ajouter = true;

                    foreach (var f in ListeFaceCorps)
                    {
                        // Si elles sont identiques, la face "faceExt" est ajoutée à la liste
                        // de face de "f"
                        if (f.FaceExtIdentique(faceExt))
                        {
                            Ajouter = false;
                            break;
                        }
                    }

                    // S'il n'y avait pas de face identique, on l'ajoute.
                    if (Ajouter)
                        ListeFaceCorps.Add(faceExt);

                }

                List<FaceGeom> ListeFaceGeom = new List<FaceGeom>();
                PlanSection = RechercherFaceProfil(ListeFaceCorps, ref ListeFaceGeom);
                ListeFaceSectionInt = TrierFacesConnectees(ListeFaceGeom);

                // Plan de la section et infos
                {
                    var v = PlanSection.Normale;
                    Double X = 0, Y = 0, Z = 0;
                    Corps.GetExtremePoint(v.X, v.Y, v.Z, out X, out Y, out Z);
                    ExtremPoint1 = new gPoint(X, Y, Z);
                    v.Inverser();
                    Corps.GetExtremePoint(v.X, v.Y, v.Z, out X, out Y, out Z);
                    ExtremPoint2 = new gPoint(X, Y, Z);
                }

                // =================================================================================

                // On recherche la face exterieure
                // s'il y a plusieurs boucles de surfaces
                if (ListeFaceSectionInt.Count > 1)
                {
                    {
                        // Si la section n'est composé que de cylindre fermé
                        Boolean EstUnCylindre = true;
                        ListFaceGeom Ext = null;
                        Double RayonMax = 0;
                        foreach (var fg in ListeFaceSectionInt)
                        {
                            if (fg.ListeFaceGeom.Count == 1)
                            {
                                var f = fg.ListeFaceGeom[0];

                                if (f.Type == eTypeFace.Cylindre)
                                {
                                    if (RayonMax < f.Rayon)
                                    {
                                        RayonMax = f.Rayon;
                                        Ext = fg;
                                    }
                                }
                                else
                                {
                                    EstUnCylindre = false;
                                    break;
                                }
                            }
                        }

                        if (EstUnCylindre)
                        {
                            FaceSectionExt = Ext;
                            ListeFaceSectionInt.Remove(Ext);
                        }
                        else
                            FaceSectionExt = null;
                    }

                    {
                        // Methode plus longue pour determiner la face exterieur
                        if (FaceSectionExt == null)
                        {
                            // On créer un vecteur perpendiculaire à l'axe du profil
                            var vect = this.PlanSection.Normale;

                            if (vect.X == 0)
                                vect = vect.Vectoriel(new gVecteur(1, 0, 0));
                            else
                                vect = vect.Vectoriel(new gVecteur(0, 0, 1));

                            vect.Normaliser();

                            // On récupère le point extreme dans cette direction
                            Double X = 0, Y = 0, Z = 0;
                            Corps.GetExtremePoint(vect.X, vect.Y, vect.Z, out X, out Y, out Z);
                            var Pt = new gPoint(X, Y, Z);

                            // La liste de face la plus proche est considérée comme la peau exterieur du profil
                            Double distMin = 1E30;
                            foreach (var Ext in ListeFaceSectionInt)
                            {
                                foreach (var fg in Ext.ListeFaceGeom)
                                {
                                    foreach (var f in fg.ListeSwFace)
                                    {
                                        Double[] res = f.GetClosestPointOn(Pt.X, Pt.Y, Pt.Z);
                                        var PtOnSurface = new gPoint(res);

                                        var dist = Pt.Distance(PtOnSurface);
                                        if (dist < 1E-6)
                                        {
                                            distMin = dist;
                                            FaceSectionExt = Ext;
                                            break;
                                        }
                                    }
                                }
                                if (FaceSectionExt.IsRef()) break;
                            }

                            // On supprime la face exterieur de la liste des faces
                            ListeFaceSectionInt.Remove(FaceSectionExt);
                        }
                    }
                }
                else
                {
                    FaceSectionExt = ListeFaceSectionInt[0];
                    ListeFaceSectionInt.RemoveAt(0);
                }
            }
            catch (Exception e) { this.LogErreur(new Object[] { e }); }

        }

        private gPlan RechercherFaceProfil(List<FaceGeom> listeFaceGeom, ref List<FaceGeom> faceExt)
        {
            gPlan? p = null;
            try
            {
                // On recherche les faces de la section
                foreach (var fg in listeFaceGeom)
                {
                    if (EstUneFaceProfil(fg))
                    {
                        faceExt.Add(fg);

                        // Si c'est un cylindre ou une extrusion, on recupère le plan
                        if ((p == null) && (fg.Type == eTypeFace.Cylindre || fg.Type == eTypeFace.Extrusion))
                            p = new gPlan(fg.Origine, fg.Direction);
                    }
                }

                // S'il n'y a que des faces plane, il faut calculer le plan de la section
                // a partir de deux plan non parallèle
                if (p == null)
                {
                    gVecteur? v1 = null;
                    foreach (var fg in faceExt)
                    {
                        if (v1 == null)
                            v1 = fg.Normale;
                        else
                        {
                            var vtmp = ((gVecteur)v1).Vectoriel(fg.Normale);
                            if (Math.Abs(vtmp.Norme) > 1E-8)
                                p = new gPlan(fg.Origine, vtmp);
                        }

                    }
                }
            }
            catch (Exception e) { this.LogErreur(new Object[] { e }); }

            return (gPlan)p;
        }

        private Boolean EstUneFaceProfil(FaceGeom fg)
        {
            foreach (var f in fg.ListeSwFace)
            {
                Byte[] Tab = Mdl.Extension.GetPersistReference3(f);
                String S = System.Text.Encoding.Default.GetString(Tab);

                int Pos_moSideFace = S.IndexOf("moSideFace3IntSurfIdRep_c");

                int Pos_moVertexRef = S.Position("moVertexRef");

                int Pos_moDerivedSurfIdRep = S.Position("moDerivedSurfIdRep_c");

                int Pos_moFromSkt = Math.Min(S.Position("moFromSktEntSurfIdRep_c"), S.Position("moFromSktEnt3IntSurfIdRep_c"));

                int Pos_moEndFace = Math.Min(S.Position("moEndFaceSurfIdRep_c"), S.Position("moEndFace3IntSurfIdRep_c"));

                if (Pos_moSideFace != -1 && Pos_moSideFace < Pos_moEndFace && Pos_moSideFace < Pos_moFromSkt && Pos_moSideFace < Pos_moVertexRef && Pos_moSideFace < Pos_moDerivedSurfIdRep)
                    return true;
            }

            return false;
        }

        private List<ListFaceGeom> TrierFacesConnectees(List<FaceGeom> listeFace)
        {
            List<FaceGeom> listeTmp = new List<FaceGeom>(listeFace);
            List<ListFaceGeom> ListeTri = null;

            if (listeTmp.Count > 0)
            {
                ListeTri = new List<ListFaceGeom>() { new ListFaceGeom(listeTmp[0]) };
                listeTmp.RemoveAt(0);

                while (listeTmp.Count > 0)
                {
                    var l = ListeTri.Last();

                    int i = 0;
                    while (i < listeTmp.Count)
                    {
                        var f = listeTmp[i];

                        if (l.AjouterFaceConnectee(f))
                        {
                            listeTmp.RemoveAt(i);
                            i = -1;
                        }
                        i++;
                    }

                    if (listeTmp.Count > 0)
                    {
                        ListeTri.Add(new ListFaceGeom(listeTmp[0]));
                        listeTmp.RemoveAt(0);
                    }
                }
            }

            // On recherche les cylindres uniques
            // et on les marque comme fermé s'ils ont plus de deux boucle
            foreach (var l in ListeTri)
            {
                if (l.ListeFaceGeom.Count == 1)
                {
                    var f = l.ListeFaceGeom[0];
                    if (f.ListeSwFace.Count == 1)
                    {
                        var cpt = 0;

                        foreach (var loop in f.SwFace.eListeDesBoucles())
                            if (loop.IsOuter()) cpt++;

                        if (cpt > 1)
                            l.Fermer = true;
                    }
                }
            }

            return ListeTri;
        }

        public enum eTypeFace
        {
            Inconnu = 1,
            Plan = 2,
            Cylindre = 3,
            Extrusion = 4
        }

        public enum eOrientation
        {
            Indefini = 1,
            Coplanaire = 2,
            Colineaire = 3,
            MemeOrigine = 4
        }

        public class FaceGeom
        {
            public Face2 SwFace = null;
            private Surface Surface = null;

            public gPoint Origine;
            public gVecteur Normale;
            public gVecteur Direction;
            public Double Rayon = 0;
            public eTypeFace Type = eTypeFace.Inconnu;

            public List<Face2> ListeSwFace = new List<Face2>();

            public List<Face2> ListeFacesConnectee
            {
                get
                {
                    var liste = new List<Face2>();

                    liste.AddRange(ListeSwFace[0].eListeDesFacesContigues());
                    for (int i = 1; i < ListeSwFace.Count; i++)
                    {
                        var l = ListeSwFace[i].eListeDesFacesContigues();

                        foreach (var f in l)
                        {
                            liste.AddIfNotExist(f);
                        }
                    }

                    return liste;
                }
            }

            public FaceGeom(Face2 swface)
            {
                SwFace = swface;

                Surface = (Surface)SwFace.GetSurface();

                ListeSwFace.Add(SwFace);

                switch ((swSurfaceTypes_e)Surface.Identity())
                {
                    case swSurfaceTypes_e.PLANE_TYPE:
                        Type = eTypeFace.Plan;
                        GetInfoPlan();
                        break;

                    case swSurfaceTypes_e.CYLINDER_TYPE:
                        Type = eTypeFace.Cylindre;
                        GetInfoCylindre();
                        break;

                    case swSurfaceTypes_e.EXTRU_TYPE:
                        Type = eTypeFace.Extrusion;
                        GetInfoExtrusion();
                        break;

                    default:
                        break;
                }
            }

            public Boolean FaceExtIdentique(FaceGeom fe, Double arrondi = 1E-10)
            {
                if (Type != fe.Type)
                    return false;

                if (!Origine.Comparer(fe.Origine, arrondi))
                    return false;

                switch (Type)
                {
                    case eTypeFace.Inconnu:
                        return false;
                    case eTypeFace.Plan:
                        if (!Normale.EstColineaire(fe.Normale, arrondi))
                            return false;
                        break;
                    case eTypeFace.Cylindre:
                        if (!Direction.EstColineaire(fe.Direction, arrondi) || (Math.Abs(Rayon - fe.Rayon) > arrondi))
                            return false;
                        break;
                    case eTypeFace.Extrusion:
                        if (!Direction.EstColineaire(fe.Direction, arrondi))
                            return false;
                        break;
                    default:
                        break;
                }

                ListeSwFace.Add(fe.SwFace);
                return true;
            }

            private void GetInfoPlan()
            {
                Boolean Reverse = SwFace.FaceInSurfaceSense();

                if (Surface.IsPlane())
                {
                    Double[] Param = Surface.PlaneParams;

                    if (Reverse)
                    {
                        Param[0] = Param[0] * -1;
                        Param[1] = Param[1] * -1;
                        Param[2] = Param[2] * -1;
                    }

                    Origine = new gPoint(Param[3], Param[4], Param[5]);
                    Normale = new gVecteur(Param[0], Param[1], Param[2]);
                }
            }

            private void GetInfoCylindre()
            {
                if (Surface.IsCylinder())
                {
                    Double[] Param = Surface.CylinderParams;

                    Origine = new gPoint(Param[0], Param[1], Param[2]);
                    Direction = new gVecteur(Param[3], Param[4], Param[5]);
                    Rayon = Param[6];

                    var UV = (Double[])SwFace.GetUVBounds();
                    Boolean Reverse = SwFace.FaceInSurfaceSense();

                    var ev1 = (Double[])Surface.Evaluate((UV[0] + UV[1]) * 0.5, (UV[2] + UV[3]) * 0.5, 0, 0);
                    if (Reverse)
                    {
                        ev1[3] = -ev1[3];
                        ev1[4] = -ev1[4];
                        ev1[5] = -ev1[5];
                    }

                    Normale = new gVecteur(ev1[3], ev1[4], ev1[5]);
                }
            }

            private void GetInfoExtrusion()
            {
                if (Surface.IsSwept())
                {
                    Double[] Param = Surface.GetExtrusionsurfParams();
                    Direction = new gVecteur(Param[0], Param[1], Param[2]);

                    Curve C = Surface.GetProfileCurve();
                    C = C.GetBaseCurve();

                    Double StartParam = 0, EndParam = 0;
                    Boolean IsClosed = false, IsPeriodic = false;

                    if (C.GetEndParams(out StartParam, out EndParam, out IsClosed, out IsPeriodic))
                    {
                        Double[] Eval = C.Evaluate(StartParam);

                        Origine = new gPoint(Eval[0], Eval[1], Eval[2]);
                    }

                    var UV = (Double[])SwFace.GetUVBounds();
                    Boolean Reverse = SwFace.FaceInSurfaceSense();

                    var ev1 = (Double[])Surface.Evaluate((UV[0] + UV[1]) * 0.5, (UV[2] + UV[3]) * 0.5, 0, 0);
                    if (Reverse)
                    {
                        ev1[3] = -ev1[3];
                        ev1[4] = -ev1[4];
                        ev1[5] = -ev1[5];
                    }

                    Normale = new gVecteur(ev1[3], ev1[4], ev1[5]);
                }
            }
        }

        public class ListFaceGeom
        {
            public Boolean Fermer = false;

            public List<FaceGeom> ListeFaceGeom = new List<FaceGeom>();

            public Double DistToExtremPoint1 = 1E30;
            public Double DistToExtremPoint2 = 1E30;

            // Initialisation avec une face
            public ListFaceGeom(FaceGeom f)
            {
                ListeFaceGeom.Add(f);
            }

            public List<Face2> ListeFaceSw()
            {
                var liste = new List<Face2>();

                foreach (var fl in ListeFaceGeom)
                    liste.AddRange(fl.ListeSwFace);

                return liste;
            }

            public Boolean AjouterFaceConnectee(FaceGeom f)
            {
                var Ajouter = false;
                var Connection = 0;

                int r = ListeFaceGeom.Count;

                for (int i = 0; i < r; i++)
                {
                    var l = ListeFaceGeom[i].ListeFacesConnectee;

                    foreach (var swf in f.ListeSwFace)
                    {
                        if (l.eContient(swf))
                        {
                            if (Ajouter == false)
                            {
                                ListeFaceGeom.Add(f);
                                Ajouter = true;
                            }

                            Connection++;
                            break;
                        }
                    }

                }

                if (Connection > 1)
                    Fermer = true;

                return Ajouter;
            }

            public void CalculerDistance(gPoint extremPoint1, gPoint extremPoint2)
            {
                foreach (var f in ListeFaceSw())
                {
                    {
                        Double[] res = f.GetClosestPointOn(extremPoint1.X, extremPoint1.Y, extremPoint1.Z);
                        var dist = extremPoint1.Distance(new gPoint(res));
                        if (dist < DistToExtremPoint1) DistToExtremPoint1 = dist;
                    }

                    {
                        Double[] res = f.GetClosestPointOn(extremPoint2.X, extremPoint2.Y, extremPoint2.Z);
                        var dist = extremPoint2.Distance(new gPoint(res));
                        if (dist < DistToExtremPoint2) DistToExtremPoint2 = dist;
                    }
                }
            }
        }

        private eOrientation Orientation(FaceGeom f1, FaceGeom f2)
        {
            var val = eOrientation.Indefini;
            if (f1.Type == eTypeFace.Plan && f2.Type == eTypeFace.Plan)
            {
                val = Orientation(f1.Origine, f1.Normale, f2.Origine, f2.Normale);
            }
            else if (f1.Type == eTypeFace.Plan && (f2.Type == eTypeFace.Cylindre || f2.Type == eTypeFace.Extrusion))
            {
                gPlan P = new gPlan(f2.Origine, f2.Direction);
                if (P.SurLePlan(f1.Origine, 1E-10) && P.SurLePlan(f1.Origine.Composer(f1.Normale), 1E-10))
                {
                    val = eOrientation.Coplanaire;
                }
            }
            else if (f2.Type == eTypeFace.Plan && (f1.Type == eTypeFace.Cylindre || f1.Type == eTypeFace.Extrusion))
            {
                gPlan P = new gPlan(f1.Origine, f1.Direction);
                if (P.SurLePlan(f2.Origine, 1E-10) && P.SurLePlan(f2.Origine.Composer(f2.Normale), 1E-10))
                {
                    val = eOrientation.Coplanaire;
                }
            }


            return val;
        }

        private eOrientation Orientation(gPoint p1, gVecteur v1, gPoint p2, gVecteur v2)
        {
            if (p1.Distance(p2) < 1E-10)
                return eOrientation.MemeOrigine;

            gVecteur Vtmp = new gVecteur(p1, p2);

            if ((v1.Vectoriel(Vtmp).Norme < 1E-10) && (v2.Vectoriel(Vtmp).Norme < 1E-10))
                return eOrientation.Colineaire;

            gVecteur Vn1 = (new gVecteur(p1, p2)).Vectoriel(v1);
            gVecteur Vn2 = (new gVecteur(p2, p1)).Vectoriel(v2);

            gVecteur Vn = Vn1.Vectoriel(Vn2);

            if (Vn.Norme < 1E-10)
                return eOrientation.Coplanaire;

            return eOrientation.Indefini;
        }

        #endregion
    }

    public class Corps : INotifyPropertyChanged
    {
        public ModelDoc2 MdlBase { get; set; }
        public Body2 SwCorps { get; set; }
        public SortedDictionary<int, int> Campagne = new SortedDictionary<int, int>();
        private int _Repere = -1;
        public int Repere
        {
            get { return _Repere; }
            set { _Repere = value; InitChemins(); }
        }

        public String RepereComplet
        {
            get { return CONSTANTES.PREFIXE_REF_DOSSIER + Repere; }
        }

        public eTypeCorps TypeCorps { get; set; }
        /// <summary>
        /// Epaisseur de la tôle ou section
        /// </summary>
        public String Dimension { get; set; }
        /// <summary>
        /// Longueur de la barre ou volume de la tôle
        /// </summary>
        public String Volume { get; set; }
        public String Materiau { get; set; }
        public ModelDoc2 Modele { get; set; }
        private long _TailleFichier = long.MaxValue;
        public String NomConfig { get; set; }
        public int IdDossier { get; set; }
        public String NomCorps { get; set; }

        public static String EnteteNomenclature(int indiceCampagne, int campagneDepartDecompte)
        {
            String entete = String.Format("{0}\t{1}", CONST_PRODUCTION.CAMPAGNE_DEPART_DECOMPTE, campagneDepartDecompte);
            entete += System.Environment.NewLine;
            entete += String.Format("{0}\t{1}\t{2}\t{3}\t{4}", "Repere", "Type", "Dimension", "Volume", "Materiau");
            for (int i = 0; i < indiceCampagne; i++)
                entete += String.Format("\t{0}", i + 1);

            return entete;
        }

        public string LigneNomenclature()
        {
            String Ligne = String.Format("{0}\t{1}\t{2}\t{3}\t{4}", Repere, TypeCorps, Dimension, Volume, Materiau);

            for (int i = 0; i < Campagne.Keys.Max(); i++)
            {
                int nb = 0;
                if (Campagne.ContainsKey(i + 1))
                    nb = Campagne[i + 1];

                Ligne += String.Format("\t{0}", nb);
            }

            return Ligne;
        }

        public static String EnteteCampagne(int indiceCampagne)
        {
            String entete = String.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}", "Repere", "Type", "Dimension", "Volume", "Materiau", indiceCampagne);
            return entete;
        }

        public int QuantiteDerniereCamapgne(int campagneDepartDecompte)
        {
            // On calcul la différence entre le total de la campagne précédente
            // et celui de la campagne actuelle
            var IndiceCampagne = Campagne.Keys.Max();
            var qteCampagneActuelle = 0;

            if (Dvp)
            {
                if (IndiceCampagne > 1)
                {
                    if (campagneDepartDecompte == IndiceCampagne)
                        qteCampagneActuelle = Math.Max(0, Campagne[IndiceCampagne]);
                    else
                        qteCampagneActuelle = Math.Max(0, Campagne[IndiceCampagne] - Campagne[IndiceCampagne - 1]);
                }
                else
                    qteCampagneActuelle = Campagne[IndiceCampagne];
            }

            return qteCampagneActuelle;
        }

        public string LigneCampagne(int campagneDepartDecompte)
        {
            var qteCampagneActuelle = QuantiteDerniereCamapgne(campagneDepartDecompte);

            String Ligne = String.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}", Repere, TypeCorps, Dimension, Volume, Materiau, qteCampagneActuelle);
            return Ligne;
        }

        public void InitCampagne(int indiceCampagne)
        {
            if (Campagne.ContainsKey(indiceCampagne))
                Campagne[indiceCampagne] = 0;
            else
                Campagne.Add(indiceCampagne, 0);
        }

        public void InitCaracteristiques(BodyFolder dossier, Body2 corps)
        {
            InitDimension(dossier, corps);
            InitVolume(dossier, corps);
        }

        private void InitDimension(BodyFolder dossier, Body2 corps)
        {
            if (TypeCorps == eTypeCorps.Tole)
                Dimension = corps.eEpaisseurCorpsOuDossier(dossier).ToString();
            else
                Dimension = dossier.eProfilDossier();
        }

        private void InitVolume(BodyFolder dossier, Body2 corps)
        {
            if (TypeCorps == eTypeCorps.Tole)
                Volume = String.Format("{0}x{1}", dossier.eLongueurToleDossier(), dossier.eLargeurToleDossier());
            else
                Volume = dossier.eLongueurProfilDossier();
        }

        public Corps(Body2 swCorps, eTypeCorps typeCorps, String materiau, ModelDoc2 mdlBase)
        {
            MdlBase = mdlBase;

            SwCorps = swCorps;
            TypeCorps = typeCorps;
            Materiau = materiau;
        }

        public Corps(String ligne, ModelDoc2 mdlBase, int indiceCampagne = 1)
        {
            MdlBase = mdlBase;

            var tab = ligne.Split(new char[] { '\t' });
            Repere = tab[0].eToInteger();
            TypeCorps = (eTypeCorps)Enum.Parse(typeof(eTypeCorps), tab[1]);
            Dimension = tab[2];
            Volume = tab[3];
            Materiau = tab[4];

            int cp = indiceCampagne;
            Campagne = new SortedDictionary<int, int>();
            for (int i = 5; i < tab.Length; i++)
                Campagne.Add(cp++, tab[i].eToInteger());
        }

        public void InitChemins()
        {
            _CheminFichierRepere = Path.Combine(MdlBase.pDossierPiece(), RepereComplet + OutilsProd.ExtPiece);
            _CheminFichierApercu = Path.Combine(MdlBase.pDossierPiece(), CONST_PRODUCTION.DOSSIER_PIECES_APERCU, RepereComplet + ".bmp");
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

        private String _CheminFichierRepere = "";
        public String CheminFichierRepere
        {
            get { return _CheminFichierRepere; }
        }

        private String _CheminFichierApercu = "";
        public String CheminFichierApercu
        {
            get { return _CheminFichierApercu; }
        }

        private BitmapImage _Apercu = null;
        public BitmapImage Apercu
        {
            get
            {
                if (_Apercu.IsNull())
                    _Apercu = (new Bitmap(_CheminFichierApercu)).ToBitmapImage();

                return _Apercu;
            }

            set { _Apercu = value; }
        }

        private Boolean _Dvp = true;
        public Boolean Dvp
        {
            get { return _Dvp; }
            set
            {
                Set(ref _Dvp, value);
            }
        }

        private String _Qte_Exp = "0";
        public String Qte_Exp
        {
            get { return _Qte_Exp; }
            set
            {
                // Si la valeur se termine par .0 on le supprime
                Regex rgx = new Regex(@"\.0$");
                value = rgx.Replace(value, "");

                // Pour eviter des mises à jour intempestives
                if (Set(ref _Qte_Exp, value) && !Maj_Qte)
                {
                    try
                    {
                        // Pour eviter des calcules intempetifs
                        Double? Eval = value.Evaluer();
                        if (Eval != null)
                            Qte = (int)Eval;
                    }
                    catch { }
                }
            }
        }

        private Boolean Maj_Qte = false;
        private int _Qte = 0;
        public int Qte
        {
            get { return _Qte; }
            set { Set(ref _Qte, value); Maj_Qte = true; Qte_Exp = value.ToString(); Maj_Qte = false; }
        }

        private String _QteSup_Exp = "0";
        public String QteSup_Exp
        {
            get { return _QteSup_Exp; }
            set
            {
                // Si la valeur se termine par .0 on le supprime
                Regex rgx = new Regex(@"\.0$");
                value = rgx.Replace(value, "");

                // Pour eviter des mises à jour intempestives
                if (Set(ref _QteSup_Exp, value))
                {
                    try
                    {
                        // Pour eviter des calcules intempetifs
                        Double? Eval = value.Evaluer();
                        if (Eval != null)
                            _QteSup = (int)Eval;
                    }
                    catch { }
                }
            }
        }

        private int _QteSup = 0;
        public int QteSup
        {
            get { return _QteSup; }
            set { _QteSup = value; _QteSup_Exp = value.ToString(); }
        }

        #region Notification WPF

        protected bool Set<U>(ref U field, U value, [CallerMemberName]string propertyName = "")
        {
            if (EqualityComparer<U>.Default.Equals(field, value)) { return false; }
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] String NomProp = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(NomProp));
        }

        #endregion
    }
}
