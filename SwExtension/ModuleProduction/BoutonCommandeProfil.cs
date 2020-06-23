using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;

namespace ModuleProduction
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Commander les profils"),
        ModuleNom("CommandeProfil")]

    public class BoutonCommandeProfil : BoutonBase
    {
        private Parametre ParamLongueurMax;
        private Double LgMax = 6000;
        private SortedDictionary<String, SortedDictionary<String, List<Double>>> DicBarres = new SortedDictionary<string, SortedDictionary<string, List<Double>>>();
        private SortedDictionary<String, SortedDictionary<String, Ligne>> DicMateriau = new SortedDictionary<string, SortedDictionary<string, Ligne>>();

        public BoutonCommandeProfil()
        {
            ParamLongueurMax = _Config.AjouterParam("LongueurMax", 6000);
        }

        protected override void Command()
        {
            try
            {
                LgMax = ParamLongueurMax.GetValeur<Double>();

                // On liste les composants
                var ListeComposants = MdlBase.pListerComposants();

                // On boucle sur les modeles
                foreach (var mdl in ListeComposants.Keys)
                {
                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                    foreach (var nomCfg in ListeComposants[mdl].Keys)
                    {
                        mdl.ShowConfiguration2(nomCfg);
                        mdl.EditRebuild3();
                        var piece = mdl.ePartDoc();
                        var ListeDossier = piece.eListeDesFonctionsDePiecesSoudees(
                            swD =>
                            {
                                BodyFolder Dossier = swD.GetSpecificFeature2();

                                // Si le dossier est la racine d'un sous-ensemble soudé, il n'y a rien dedans
                                if (Dossier.IsRef() && (Dossier.eEstExclu() == false) && (Dossier.eNbCorps() > 0) &&
                                (eTypeCorps.Barre).HasFlag(Dossier.eTypeDeDossier()))
                                    return true;

                                return false;
                            }
                            );

                        var NbConfig = ListeComposants[mdl][nomCfg];

                        foreach (var fDossier in ListeDossier)
                        {
                            BodyFolder Dossier = fDossier.GetSpecificFeature2();
                            var SwCorps = Dossier.ePremierCorps();
                            var NomCorps = SwCorps.Name;
                            var MateriauCorps = SwCorps.eGetMateriauCorpsOuPiece(piece, nomCfg);
                            var NbCorps = Dossier.eNbCorps() * NbConfig;
                            var Profil = Dossier.eProfilDossier();
                            var Longueur = Dossier.eLongueurProfilDossier().eToDouble();

                            for (int i = 0; i < NbCorps; i++)
                            {
                                var LgTmp = Longueur;

                                while (LgTmp > LgMax)
                                {
                                    RassemblerBarre(MateriauCorps, Profil, LgMax);
                                    LgTmp = LgTmp - LgMax;
                                }

                                RassemblerBarre(MateriauCorps, Profil, LgTmp);
                            }
                        }
                    }

                    mdl.eFermerSiDifferent(MdlBase);
                }

                // On fait la mise en barre
                foreach (var Materiau in DicBarres.Keys)
                {
                    WindowLog.SautDeLigne();
                    WindowLog.EcrireF("{0}", Materiau); ;
                    foreach (var Profil in DicBarres[Materiau].Keys)
                    {
                        WindowLog.SautDeLigne();
                        WindowLog.EcrireF("  {0}", Profil);
                        var lst = DicBarres[Materiau][Profil];
                        lst.Sort();
                        lst.Reverse();

                        foreach (var lg in lst)
                            WindowLog.EcrireF("  - {0}", lg);

                        int i = 0;
                        while (true)
                        {
                            var lg = lst[i];
                            if (AjouterBarre(Materiau, Profil, lg))
                            {
                                lst.RemoveAt(i);
                                if (lst.Count == 0) break;
                            }
                            else
                                i++;

                            if (i >= lst.Count)
                            {
                                i = 0;
                                NouvelleBarre(Materiau, Profil);
                            }
                        }
                    }
                }

                String Resume = "";
                foreach (var Materiau in DicMateriau.Keys)
                {
                    Resume += System.Environment.NewLine;
                    Resume += String.Format("{0}", Materiau);
                    foreach (var Profil in DicMateriau[Materiau].Keys)
                    {
                        var Ligne = DicMateriau[Materiau][Profil];
                        Resume += System.Environment.NewLine;
                        Resume += String.Format("   {1,4}×  {0,-25}  [{2:N0}]", Profil, Ligne.NbBarre, Ligne.Reste);
                    }
                }
                WindowLog.SautDeLigne();
                WindowLog.Ecrire("A commander");
                WindowLog.Ecrire(Resume);
                WindowLog.SautDeLigne();
                File.WriteAllText(Path.Combine(MdlBase.eDossier(), "CommandeProfil.txt"), Resume);
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        private void RassemblerBarre(String materiauCorps, String profil, Double longueur)
        {
            if (DicBarres.ContainsKey(materiauCorps))
            {
                if (DicBarres[materiauCorps].ContainsKey(profil))
                {
                    var l = DicBarres[materiauCorps][profil];
                    l.Add(longueur);
                }
                else
                {
                    var l = new List<Double>();
                    l.Add(longueur);
                    DicBarres[materiauCorps].Add(profil, l);
                }
            }
            else
            {
                var l = new List<Double>();
                l.Add(longueur);
                var DicProfil = new SortedDictionary<String, List<Double>>();
                DicProfil.Add(profil, l);
                DicBarres.Add(materiauCorps, DicProfil);
            }
        }

        private void NouvelleBarre(String materiauCorps, String profil)
        {
            if (DicMateriau.ContainsKey(materiauCorps))
            {
                if (DicMateriau[materiauCorps].ContainsKey(profil))
                {
                    var l = DicMateriau[materiauCorps][profil];
                    l.NbBarre++;
                    l.Reste = LgMax;
                }
            }
        }

        private Boolean AjouterBarre(String materiauCorps, String profil, Double longueur)
        {
            Boolean AjouterBarre = true;
            if (DicMateriau.ContainsKey(materiauCorps))
            {
                if (DicMateriau[materiauCorps].ContainsKey(profil))
                {
                    var l = DicMateriau[materiauCorps][profil];
                    if (longueur > l.Reste)
                        AjouterBarre = false;
                    else
                        l.Reste -= longueur;
                }
                else
                {
                    var l = new Ligne();
                    l.NbBarre = 1;
                    l.Reste = LgMax - longueur;
                    DicMateriau[materiauCorps].Add(profil, l);
                }
            }
            else
            {
                var l = new Ligne();
                l.NbBarre = 1;
                l.Reste = LgMax - longueur;
                var DicProfil = new SortedDictionary<string, Ligne>();
                DicProfil.Add(profil, l);
                DicMateriau.Add(materiauCorps, DicProfil);
            }

            return AjouterBarre;
        }

        private class Ligne
        {
            public int NbBarre = 0;
            public Double Reste = 0;
        }
    }
}
