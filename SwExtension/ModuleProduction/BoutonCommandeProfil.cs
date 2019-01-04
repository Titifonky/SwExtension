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
        private ModelDoc2 MdlBase = null;
        private Parametre ParamLongueurMax;
        private Double LgMax = 6000;
        private SortedDictionary<String, SortedDictionary<String, Ligne>> DicMateriau = new SortedDictionary<string, SortedDictionary<string, Ligne>>();

        public BoutonCommandeProfil()
        {
            ParamLongueurMax = _Config.AjouterParam("LongueurMax", 6000);
        }

        protected override void Command()
        {
            try
            {
                MdlBase = App.ModelDoc2;
                LgMax = ParamLongueurMax.GetValeur<Double>();

                // On liste les composants
                var ListeComposants = MdlBase.pListerComposants(false);

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
                                if (Dossier.IsRef() && (Dossier.eNbCorps() > 0) &&
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
                                    AjouterBarre(MateriauCorps, Profil, LgMax);
                                    LgTmp = LgTmp - LgMax;
                                }

                                AjouterBarre(MateriauCorps, Profil, LgTmp);
                            }
                        }
                    }

                    if (mdl.GetPathName() != MdlBase.GetPathName())
                        App.Sw.CloseDoc(mdl.GetPathName());
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
                WindowLog.Ecrire("Liste");
                WindowLog.Ecrire(Resume);
                WindowLog.SautDeLigne();
                File.WriteAllText(Path.Combine(MdlBase.eDossier(), "CommandeProfil.txt"), Resume);
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        private void AjouterBarre(String materiauCorps, String profil, Double longueur)
        {
            if (DicMateriau.ContainsKey(materiauCorps))
            {
                if (DicMateriau[materiauCorps].ContainsKey(profil))
                {
                    var l = DicMateriau[materiauCorps][profil];
                    if (l.Reste < longueur)
                    {
                        l.NbBarre += 1;
                        l.Reste = LgMax - longueur;
                    }
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
        }

        private class Ligne
        {
            public int NbBarre = 0;
            public Double Reste = 0;
        }
    }
}
