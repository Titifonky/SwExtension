using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Lister les materiaux"),
        ModuleNom("ListerMateriaux")]
    public class BoutonListerMateriaux : BoutonBase
    {
        private class Info
        {
            public String Key { get { return Materiau + "_" + Base; } }
            public String Classe { get; private set; }
            public String Materiau { get; private set; }
            public String Base { get; private set; }
            public Double Volume { get; set; }
            public Double Densite { get; private set; }
            public Double Masse { get; private set; }
            public Boolean Herite { get; set; }

            private static Dictionary<String, Double> _DicDensite = new Dictionary<String, Double>();
            private static Dictionary<String, String> _DicClasse = new Dictionary<String, String>();

            public Info(String materiau, String baseMat, Double volume, Boolean herite = false)
            {
                Materiau = materiau;
                Base = baseMat;
                Volume = volume;
                Herite = herite;

                if (_DicDensite.ContainsKey(Materiau))
                {
                    Densite = _DicDensite[Materiau];
                    Classe = _DicClasse[Materiau];
                }
                else
                {
                    String _classe;
                    Densite = Sw.eProprieteMat(Base, Materiau, eMatPropriete.densite, out _classe);
                    _DicDensite.Add(Materiau, Densite);
                    Classe = _classe;
                    _DicClasse.Add(Materiau, _classe);
                }

                Masse = Volume * Densite;
            }

            public void Ajouter(Info info)
            {
                    Masse += info.Masse;
                    Volume += info.Volume;
            }

            public Info Clone()
            {
                return new Info(Materiau, Base, Volume, Herite);
            }

            public void Multiplier(int nb)
            {
                Volume = Volume * nb;
                Masse = Masse * nb;
            }
        }

        private class Piece
        {
            public Component2 Composant { get; private set; }
            public int Nb { get; set; }
            public Info Info { get; private set; }
            public Double Masse { get; private set; }
            private Dictionary<String, Dictionary<String, Info>> _DicClasse = new Dictionary<String, Dictionary<String, Info>>();
            public Dictionary<String, Dictionary<String, Info>> ListeClasse { get { return _DicClasse; } }

            private static Dictionary<String, Double> _DicDensite = new Dictionary<String, Double>();

            public Piece(Component2 cp)
            {
                Composant = cp;
                Nb = 1;
                String Base = "";
                String Materiau = cp.eGetMateriau(out Base);
                Info = new Info(Materiau, Base, 0);
            }

            public Info AjouterCorps(Body2 body, int nb)
            {
                Info inf = null;

                for (int i = 0; i < nb; i++)
                    inf = AjouterCorps(body);

                return inf;
            }

            public Info AjouterCorps(Body2 body)
            {
                String Base = null;
                String Materiau = body.eGetMateriau(Composant.eNomConfiguration(), out Base);
                Boolean Herite = false;

                if (String.IsNullOrWhiteSpace(Materiau))
                {
                    Base = Info.Base;
                    Materiau = Info.Materiau;
                    Herite = true;
                }

                Info infoMat = new Info(Materiau, Base, body.eVolume(), Herite);

                if (_DicClasse.ContainsKey(infoMat.Classe))
                {
                    Dictionary<String, Info> _DcMateriaux = _DicClasse[infoMat.Classe];
                    if (_DcMateriaux.ContainsKey(infoMat.Key))
                        _DcMateriaux[infoMat.Key].Ajouter(infoMat);
                    else
                        _DcMateriaux.Add(infoMat.Key, infoMat);
                }
                else
                {
                    Dictionary<String, Info> _DcMateriaux = new Dictionary<String, Info>();
                    _DcMateriaux.Add(infoMat.Key, infoMat);
                    _DicClasse.Add(infoMat.Classe, _DcMateriaux);
                }

                Masse += infoMat.Masse;

                return infoMat;
            }
        }

        private class Ensemble
        {
            public Double Masse { get; private set; }
            private Dictionary<String, Dictionary<String, Info>> _DicClasse = new Dictionary<String, Dictionary<String, Info>>();
            public Dictionary<String, Dictionary<String, Info>> ListeClasse { get { return _DicClasse; } }

            public Ensemble() { }

            public void AjouterPiece(Piece piece)
            {
                Masse += piece.Masse * piece.Nb;

                foreach (String classe in piece.ListeClasse.Keys)
                {
                    Dictionary<String, Info> _DcMateriaux;
                    if (_DicClasse.ContainsKey(classe))
                    {
                        _DcMateriaux = _DicClasse[classe];
                    }
                    else
                    {
                        _DcMateriaux = new Dictionary<String, Info>();
                        _DicClasse.Add(classe, _DcMateriaux);
                    }

                    foreach (String materiau in piece.ListeClasse[classe].Keys)
                    {
                        Info inf = piece.ListeClasse[classe][materiau].Clone();
                        inf.Multiplier(piece.Nb);

                        if (_DcMateriaux.ContainsKey(materiau))
                            _DcMateriaux[materiau].Ajouter(inf);
                        else
                            _DcMateriaux.Add(materiau, inf);
                    }
                }
            }
        }

        /// <summary>
        /// Liste les corps avec Composant.eListeDesDossiersDePiecesSoudees().eListeDesCorps()
        /// Permet de regrouper
        /// </summary>
        protected override void Command()
        {
            try
            {
                Dictionary<String, Piece> _DicPiece = new Dictionary<String, Piece>();
                int NbTotalComposant = 0;

                // Si c'est une piece, on ajoute le composant racine
                if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                {
                    _DicPiece.Add(MdlBase.eComposantRacine().eKeyAvecConfig(), new Piece(MdlBase.eComposantRacine()));
                    NbTotalComposant = 1;
                }

                // Si c'est un assemblage, on liste les composants
                if (MdlBase.TypeDoc() == eTypeDoc.Assemblage)
                    MdlBase.eRecListeComposant(
                        c =>
                        {
                            if (!c.IsHidden(false) && (c.TypeDoc() == eTypeDoc.Piece))
                            {
                                if (_DicPiece.ContainsKey(c.eKeyAvecConfig()))
                                {
                                    _DicPiece[c.eKeyAvecConfig()].Nb += 1;
                                    return true;
                                }

                                _DicPiece.Add(c.eKeyAvecConfig(), new Piece(c));
                                NbTotalComposant++;
                                return true;
                            }
                            return false;
                        },
                        null,
                        true);
                   

                WindowLog.Ecrire("Nb de composants : " + NbTotalComposant);
                WindowLog.SautDeLigne();

                List<Piece> ListePiece = _DicPiece.Values.ToList<Piece>();
                
                // On trie
                ListePiece.Sort((c1, c2) => { WindowsStringComparer sc = new WindowsStringComparer(); return sc.Compare(c1.Composant.eKeyAvecConfig(), c2.Composant.eKeyAvecConfig()); });

                Ensemble Ensemble = new Ensemble();

                foreach (Piece piece in ListePiece)
                {
                    //WindowLog.Ecrire(piece.Nb + " × " + piece.Composant.eNomSansExt() + " Config \"" + piece.Composant.eNomConfiguration() + "\"");
                    //WindowLog.Ecrire(" -> Mat : " + piece.Info.Materiau + " / Base : " + piece.Info.Base);

                    var ListeDossier = piece.Composant.eListeDesDossiersDePiecesSoudees();

                    if (piece.Composant.IsRoot())
                        ListeDossier = piece.Composant.ePartDoc().eListeDesDossiersDePiecesSoudees();

                    foreach (var dossier in ListeDossier)
                    {
                        if (dossier.eNbCorps() == 0) continue;

                        var listeCorps = dossier.eListeDesCorps();

                        if (listeCorps.IsNull()) continue;

                        Body2 corps = listeCorps[0];

                        Info InfoCorps = piece.AjouterCorps(corps, dossier.eNbCorps());

                        String herite = InfoCorps.Herite ? " (herité) " : "";

                        String nomCorps = corps.Name;
                        if (!dossier.eNom().eIsLike(CONSTANTES.ARTICLE_LISTE_DES_PIECES_SOUDEES, false))
                            nomCorps = dossier.eNom();

                        String Ligne = String.Format(" -- ×{5} {0} -> {1} : {2:0.0} kg / Base : {3}{4}",
                            nomCorps,
                            InfoCorps.Materiau,
                            InfoCorps.Masse,
                            InfoCorps.Base,
                            herite,
                            dossier.eNbCorps()
                            );

                        //WindowLog.Ecrire(Ligne);
                    }

                    Ensemble.AjouterPiece(piece);

                    if (piece.Masse > 0.01)
                    {
                        List<String> TextePiece = new List<String>();
                        foreach (String classe in piece.ListeClasse.Keys)
                        {
                            Double Masse = 0;

                            foreach (Info i in piece.ListeClasse[classe].Values)
                            {
                                Masse += i.Masse;
                                TextePiece.Add(String.Format("  -- {0} : {1:0.0} kg", i.Materiau, i.Masse));
                            }
                            TextePiece.Insert(0, String.Format("  {0} : {1:0.0} kg", classe, Masse));
                        }
                        TextePiece.Insert(0, String.Format("Poids de la piece : {0:0.0} kg", piece.Masse));
                        //WindowLog.Ecrire(TextePiece);
                    }

                    //WindowLog.SautDeLigne();
                }

                WindowLog.SautDeLigne();

                WindowLog.Ecrire("Resultat :");
                WindowLog.Ecrire("----------------");

                List<String> TexteEnsemble = new List<String>();
                foreach (String classe in Ensemble.ListeClasse.Keys)
                {
                    Double Masse = 0;

                    List<String> Texte = new List<String>();
                    foreach (Info i in Ensemble.ListeClasse[classe].Values)
                    {
                        Masse += i.Masse;
                        Texte.Add(String.Format("  -- {0} : {1:0.0} kg", i.Materiau, i.Masse));
                    }

                    if (Masse > 0.01)
                    {
                        TexteEnsemble.Add(String.Format("  {0} : {1:0.0} kg", classe, Masse));
                        TexteEnsemble.AddRange(Texte);
                    }
                }
                TexteEnsemble.Insert(0, String.Format("Poids de l'ensemble : {0:0.0} kg", Ensemble.Masse));
                WindowLog.Ecrire(TexteEnsemble);
                WindowLog.SautDeLigne();

            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        /// <summary>
        /// Liste les corps avec Composant.eListeDesCorps()
        /// </summary>
        protected void OldCommand()
        {
            try
            {
                Dictionary<String, Piece> _DicPiece = new Dictionary<String, Piece>();
                int NbTotalComposant = 0;

                // Si c'est une piece, on ajoute le composant racine
                if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                {
                    _DicPiece.Add(MdlBase.eComposantRacine().eKeyAvecConfig(), new Piece(MdlBase.eComposantRacine()));
                    NbTotalComposant = 1;
                }

                // Si c'est un assemblage, on liste les composants
                if (MdlBase.TypeDoc() == eTypeDoc.Assemblage)
                    MdlBase.eRecListeComposant(
                        c =>
                        {
                            if (!c.IsHidden(false) && (c.TypeDoc() == eTypeDoc.Piece))
                            {
                                if (_DicPiece.ContainsKey(c.eKeyAvecConfig()))
                                {
                                    _DicPiece[c.eKeyAvecConfig()].Nb += 1;
                                    return true;
                                }

                                _DicPiece.Add(c.eKeyAvecConfig(), new Piece(c));
                                NbTotalComposant++;
                                return true;
                            }
                            return false;
                        },
                        null,
                        true);


                WindowLog.Ecrire("Nb de composants : " + NbTotalComposant);
                WindowLog.SautDeLigne();

                List<Piece> ListePiece = _DicPiece.Values.ToList<Piece>();

                // On trie
                ListePiece.Sort((c1, c2) => { WindowsStringComparer sc = new WindowsStringComparer(); return sc.Compare(c1.Composant.eKeyAvecConfig(), c2.Composant.eKeyAvecConfig()); });

                Ensemble Ensemble = new Ensemble();

                foreach (Piece piece in ListePiece)
                {
                    WindowLog.Ecrire(piece.Nb + " × " + piece.Composant.eNomAvecExt() + " Config \"" + piece.Composant.eNomConfiguration() + "\"");
                    WindowLog.Ecrire(" -> Mat : " + piece.Info.Materiau + " / Base : " + piece.Info.Base);

                    foreach (Body2 corps in piece.Composant.eListeCorps())
                    {
                        Info InfoCorps = piece.AjouterCorps(corps);

                        String herite = InfoCorps.Herite ? " (herité) " : "";

                        String Ligne = String.Format(" -- {0} -> {1} : {2:0.0} kg / Base : {3}{4}",
                            corps.Name,
                            InfoCorps.Materiau,
                            InfoCorps.Masse,
                            InfoCorps.Base,
                            herite
                            );

                        WindowLog.Ecrire(Ligne);
                    }

                    Ensemble.AjouterPiece(piece);

                    if (piece.Masse > 0.01)
                    {
                        List<String> TextePiece = new List<String>();
                        foreach (String classe in piece.ListeClasse.Keys)
                        {
                            Double Masse = 0;

                            foreach (Info i in piece.ListeClasse[classe].Values)
                            {
                                Masse += i.Masse;
                                TextePiece.Add(String.Format("  -- {0} : {1:0.0} kg", i.Materiau, i.Masse));
                            }
                            TextePiece.Insert(0, String.Format("  {0} : {1:0.0} kg", classe, Masse));
                        }
                        TextePiece.Insert(0, String.Format("Poids de la piece : {0:0.0} kg", piece.Masse));
                        WindowLog.Ecrire(TextePiece);
                        WindowLog.SautDeLigne();
                    }
                }

                WindowLog.SautDeLigne();

                WindowLog.Ecrire("Resultat :");
                WindowLog.Ecrire("----------------");

                List<String> TexteEnsemble = new List<String>();
                foreach (String classe in Ensemble.ListeClasse.Keys)
                {
                    Double Masse = 0;

                    List<String> Texte = new List<String>();
                    foreach (Info i in Ensemble.ListeClasse[classe].Values)
                    {
                        Masse += i.Masse;
                        Texte.Add(String.Format("  -- {0} : {1:0.0} kg", i.Materiau, i.Masse));
                    }

                    if (Masse > 0.01)
                    {
                        TexteEnsemble.Add(String.Format("  {0} : {1:0.0} kg", classe, Masse));
                        TexteEnsemble.AddRange(Texte);
                    }
                }
                TexteEnsemble.Insert(0, String.Format("Poids de l'ensemble : {0:0.0} kg", Ensemble.Masse));
                WindowLog.Ecrire(TexteEnsemble);
                WindowLog.SautDeLigne();

            }
            catch (Exception e)
            { Log.Message(e); }
        }
    }
}
