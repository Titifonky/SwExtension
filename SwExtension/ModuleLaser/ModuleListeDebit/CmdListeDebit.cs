using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;

namespace ModuleLaser
{
    namespace ModuleListeDebit
    {
        public enum eTypeSortie
        {
            ListeDebit = 1,
            ListeBarre = 2
        }

        public class CmdListeDebit : Cmd
        {
            public ModelDoc2 MdlBase = null;
            public List<String> ListeMateriaux = new List<String>();
            public int Quantite = 1;
            public Boolean PrendreEnCompteTole = false;
            public Boolean ComposantsExterne = false;
            public String RefFichier = "";

            public eTypeSortie TypeSortie = eTypeSortie.ListeDebit;
            public String ForcerMateriau = null;

            public String CheminFichier { get; private set; }

            public int LgBarre = 6000;

            private ListeElement listeElement;

            public ListeLgProfil Analyser()
            {
                InitTime();

                try
                {
                    listeElement = new ListeElement(LgBarre);

                    eTypeCorps Filtre = PrendreEnCompteTole ? eTypeCorps.Barre | eTypeCorps.Tole : eTypeCorps.Barre;
                    HashSet<String> HashMateriaux = new HashSet<string>(ListeMateriaux);

                    var dic = MdlBase.DenombrerDossiers(ComposantsExterne, HashMateriaux, Filtre);

                    foreach (var mdl in dic.Keys)
                    {
                        mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                        foreach (var NomConfigPliee in dic[mdl].Keys)
                        {
                            mdl.ShowConfiguration2(NomConfigPliee);
                            mdl.EditRebuild3();
                            PartDoc Piece = mdl.ePartDoc();

                            var ListeDossier = dic[mdl][NomConfigPliee];
                            foreach (var t in ListeDossier)
                            {
                                var IdDossier = t.Key;
                                var QuantiteBarre = t.Value * Quantite;

                                Feature fDossier = Piece.FeatureById(IdDossier);
                                BodyFolder dossier = fDossier.GetSpecificFeature2();
                                var RefDossier = dossier.eProp(CONSTANTES.REF_DOSSIER);

                                String Profil = dossier.eProp(CONSTANTES.PROFIL_NOM);
                                Double Longueur = dossier.eProp(CONSTANTES.PROFIL_LONGUEUR).eToDouble();
                                Double A = dossier.eProp(CONSTANTES.PROFIL_ANGLE1).eToDouble();
                                Double B = dossier.eProp(CONSTANTES.PROFIL_ANGLE2).eToDouble();

                                if (String.IsNullOrWhiteSpace(Profil) || (Longueur == 0)) continue;

                                var Materiau = dossier.ePremierCorps().eGetMateriauCorpsOuPiece(Piece, NomConfigPliee);

                                if (!HashMateriaux.Contains(Materiau)) continue;

                                Materiau = ForcerMateriau.IsRefAndNotEmpty(Materiau);

                                listeElement.AjouterElement(QuantiteBarre, RefDossier, Materiau, Profil, Longueur, A, B);
                            }
                        }
                    }

                    ExecuterEn();
                    return listeElement.ListeLgProfil;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }

                return null;
            }

            protected override void Command()
            {
                try
                {
                    switch (TypeSortie)
                    {
                        case eTypeSortie.ListeDebit:
                            {
                                var DicBarre = listeElement.ListeBarre();

                                String ResumeBarre = DicBarre.ResumeNbBarre();
                                String ResumeListeDebit = DicBarre.ResumeListeDebit();

                                String Complet = RefFichier + "\r\n" + ResumeBarre + "\r\n\r\n" + ResumeListeDebit;

                                WindowLog.Ecrire(ResumeBarre);

                                CheminFichier = Path.Combine(MdlBase.eDossier(), RefFichier + " - " + MdlBase.eNomSansExt() + " - Liste de d�bit.txt");

                                StreamWriter s = new StreamWriter(CheminFichier);
                                s.Write(Complet);
                                s.Close();

                                WindowLog.SautDeLigne(2);
                                WindowLog.EcrireF("Nb d'�l�ments {0}", listeElement.NbElement);
                                WindowLog.EcrireF("Nb de barres {0}", DicBarre.NbBarre);
                            }
                            break;
                        case eTypeSortie.ListeBarre:
                            {
                                String ResumeListeBarre = listeElement.ResumeListeBarre();

                                CheminFichier = Path.Combine(MdlBase.eDossier(), RefFichier + " - " + MdlBase.eNomSansExt() + " - Liste des barres.txt");
                                String Complet = RefFichier + "\r\n" + ResumeListeBarre;

                                WindowLog.Ecrire(ResumeListeBarre);

                                StreamWriter s = new StreamWriter(CheminFichier);
                                s.Write(Complet);
                                s.Close();

                                WindowLog.SautDeLigne(2);
                            }
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            public String ConstruireRefBarre(ModelDoc2 mdl, String configPliee, String noDossier)
            {
                return String.Format("{0}-{1}-{2}", mdl.eNomSansExt(), configPliee, noDossier);
            }

            private class Element
            {
                public String Ref { get; private set; }
                public String Materiau { get; private set; }
                public String Profil { get; private set; }
                public Double Lg { get; private set; }
                public Double A { get; private set; }
                public Double B { get; private set; }

                public Element(String reference, String materiau, String profil, Double lg, Double a, Double b)
                {
                    Ref = reference; Materiau = materiau; Profil = profil; Lg = lg; A = a; B = b;
                }
            }

            private class Barre : List<Element>
            {
                public Double LgTT { get; private set; }
                public Double LgUtilise { get; private set; }
                public Double LgChute { get; private set; }

                public Barre(Double lgMax)
                {
                    LgTT = lgMax;
                    LgUtilise = 0;
                    LgChute = LgTT;
                }

                public Barre(Double lgMax, Element e)
                {
                    LgTT = lgMax;
                    LgUtilise = 0;
                    LgChute = LgTT;

                    Ajouter(e);
                }

                public Boolean Ajouter(Element e)
                {
                    if ((e.Lg <= LgChute) || (e.Lg > LgTT))
                    {
                        Add(e);
                        LgUtilise += e.Lg;
                        LgChute = LgTT - LgUtilise;
                        return true;
                    }

                    return false;
                }
            }

            public class ListeLgProfil
            {
                private Double _LgBarre = 6000;
                private Dictionary<String, Dictionary<String, Double>> _dic = new Dictionary<string, Dictionary<string, Double>>();

                public Dictionary<String, Dictionary<String, Double>> DicLg { get { return _dic; } }

                public ListeLgProfil(int lgBarre)
                {
                    _LgBarre = lgBarre;
                }

                public void AjouterProfil(String materiau, String profil)
                {
                    var DicMat = new Dictionary<String, Double>();

                    if (_dic.ContainsKey(materiau))
                        DicMat = _dic[materiau];
                    else
                        _dic.Add(materiau, DicMat);

                    if (!DicMat.ContainsKey(profil))
                    {
                        DicMat.Add(profil, 6000);
                        NbProfils++;
                    }
                }

                public Double LgMax(String materiau, String profil)
                {
                    Double LgMax = 0;

                    if (_dic.ContainsKey(materiau))
                    {
                        if (_dic[materiau].ContainsKey(profil))
                            LgMax = _dic[materiau][profil];
                    }

                    return LgMax;
                }

                public int NbProfils { get; private set; }
            }

            private class ListeElement
            {
                private Dictionary<String, Dictionary<String, List<Element>>> _dic = new Dictionary<string, Dictionary<string, List<Element>>>();

                private ListeLgProfil _LgProfils;

                public ListeLgProfil ListeLgProfil { get { return _LgProfils; } }

                public ListeElement(int lgBarre)
                {
                    _LgProfils = new ListeLgProfil(lgBarre);
                }

                public ListeBarre ListeBarre()
                {
                    Trier();

                    ListeBarre ListeBarre = new ListeBarre(_LgProfils);

                    foreach (var Materiau in _dic.Values)
                    {
                        foreach (var ListProf in Materiau.Values)
                        {
                            foreach (var Element in ListProf)
                                ListeBarre.AjouterElement(Element);
                        }
                    }

                    return ListeBarre;
                }

                public void AjouterElement(int Nb, String reference, String materiau, String profil, Double lg, Double a, Double b)
                {
                    var DicMat = new Dictionary<String, List<Element>>();

                    if (_dic.ContainsKey(materiau))
                        DicMat = _dic[materiau];
                    else
                        _dic.Add(materiau, DicMat);

                    var ListProf = new List<Element>();

                    if (DicMat.ContainsKey(profil))
                        ListProf = DicMat[profil];
                    else
                        DicMat.Add(profil, ListProf);

                    for (int i = 0; i < Nb; i++)
                    {
                        NbElement++;
                        ListProf.Add(new Element(reference, materiau, profil, lg, a, b));
                    }

                    _LgProfils.AjouterProfil(materiau, profil);
                }

                private void Trier()
                {
                    foreach (var Materiau in _dic.Values)
                    {
                        foreach (var ListProf in Materiau.Values)
                        {
                            ListProf.Sort(
                                (e1, e2) =>
                                {
                                    if (e1.Lg == e2.Lg) return 0;
                                    if (e1.Lg > e2.Lg) return -1;
                                    return 1;
                                });
                        }
                    }
                }

                public String ResumeListeBarre()
                {
                    String text = "";

                    foreach (var Materiau in _dic.Keys)
                    {
                        text += "\r\n" + Materiau;
                        var DicMat = _dic[Materiau];

                        foreach (var Prof in DicMat.Keys)
                        {
                            var ListProf = DicMat[Prof];

                            int nbEle = 0;
                            double Lg = ListProf[0].Lg;

                            String format = "\r\n  {0,-30} {1,8:0.0}mm {2,5}";

                            foreach (var ele in ListProf)
                            {
                                if (ele.Lg != Lg)
                                {
                                    text += String.Format(format, Prof, Lg, "�" + nbEle);
                                    nbEle = 1;
                                    Lg = ele.Lg;
                                }
                                else
                                    nbEle += 1;
                            }

                            text += String.Format(format, Prof, Lg, "�" + nbEle);
                        }
                    }

                    return text;
                }

                public int NbElement { get; private set; }
            }

            private class ListeBarre
            {
                private Dictionary<String, Dictionary<String, List<Barre>>> _dic = new Dictionary<String, Dictionary<String, List<Barre>>>();

                private ListeLgProfil _ListeLgProfil;

                public ListeBarre(ListeLgProfil listeLgProfil)
                {
                    _ListeLgProfil = listeLgProfil;
                }

                public void AjouterElement(Element element)
                {
                    var DicMat = new Dictionary<String, List<Barre>>();

                    if (_dic.ContainsKey(element.Materiau))
                        DicMat = _dic[element.Materiau];
                    else
                        _dic.Add(element.Materiau, DicMat);

                    var ListBarre = new List<Barre>();

                    if (DicMat.ContainsKey(element.Profil))
                        ListBarre = DicMat[element.Profil];
                    else
                        DicMat.Add(element.Profil, ListBarre);

                    Double lgMax = _ListeLgProfil.LgMax(element.Materiau, element.Profil);
                    
                    if(ListBarre.Count == 0)
                    {
                        NbBarre++;
                        ListBarre.Add(new Barre(lgMax));
                    }

                    // On essaye d'ajouter la barre
                    // si c'est ok, on sort de la routine
                    foreach (var barre in ListBarre)
                    {
                        if (barre.Ajouter(element))
                            return;
                    }

                    // sinon on ajoute une nouvelle barre
                    {
                        var barre = new Barre(lgMax);
                        barre.Ajouter(element);
                        NbBarre++;
                        ListBarre.Add(barre);
                    }
                    
                }

                public String ResumeNbBarre()
                {
                    String text = "";

                    foreach (var Materiau in _dic.Keys)
                    {
                        text += "\r\n" + Materiau;
                        var DicMat = _dic[Materiau];
                        
                        foreach (var Prof in DicMat.Keys)
                        {
                            var ListBarre = DicMat[Prof];
                            text += String.Format("\r\n  {0,5}� {1}", ListBarre.Count, Prof);
                        }
                    }

                    return text;
                }

                public String ResumeListeDebit()
                {
                    String text = "";

                    foreach (var Materiau in _dic.Keys)
                    {
                        text += "\r\n" + Materiau;
                        var DicMat = _dic[Materiau];

                        foreach (var Prof in DicMat.Keys)
                        {
                            var ListBarre = DicMat[Prof];
                            text += String.Format("\r\n {0,4}� {1}", ListBarre.Count, Prof);

                            for (int i = 0; i < ListBarre.Count; i++)
                            {
                                var barre = ListBarre[i];
                                text += String.Format("\r\n      {3}Barre {0,-3} : {1,4:0}mm \\ Chute : {2,4:0}mm", i+1, barre.LgUtilise, barre.LgChute, barre.LgChute < 0 ? "!!!!!!! " : "");

                                foreach (var ele in barre)
                                {
                                    text += String.Format("\r\n         {0}   {1,6:0.0}mm   {2,4:0.0}�   {3,4:0.0}�", ele.Ref, ele.Lg, ele.A, ele.B);
                                }

                                text += "\r\n";
                            }
                            text += "\r\n";
                        }
                    }
                    return text;
                }

                public int NbBarre { get; private set; }
            }
        }
    }
}


