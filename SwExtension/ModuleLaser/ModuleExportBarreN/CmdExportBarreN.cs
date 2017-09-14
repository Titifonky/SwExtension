using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ModuleLaser
{
    namespace ModuleExportBarreN
    {
        public class CmdExportBarreN : Cmd
        {
            public ModelDoc2 MdlBase = null;
            public List<String> ListeMateriaux = new List<String>();
            public int Quantite = 1;
            public Boolean PrendreEnCompteTole = false;
            public Boolean ComposantsExterne = false;
            public String RefFichier = "";

            public Boolean ReinitialiserNoDossier = false;
            public Boolean MajListePiecesSoudees = false;
            public String ForcerMateriau = null;

            private HashSet<String> HashMateriaux;
            private Dictionary<String, int> DicQte = new Dictionary<string, int>();
            private Dictionary<String, List<String>> DicConfig = new Dictionary<String, List<String>>();
            private List<Component2> ListeCp = new List<Component2>();

            private ListeProfil listeProfil;

            public ListeProfil Analyser()
            {
                InitTime();

                try
                {
                    listeProfil = new ListeProfil();

                    HashMateriaux = new HashSet<string>(ListeMateriaux);

                    if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                    {
                        Component2 CpRacine = MdlBase.eComposantRacine();
                        ListeCp.Add(CpRacine);

                        if (!CpRacine.eNomConfiguration().eEstConfigPliee())
                        {
                            WindowLog.Ecrire("Pas de configuration valide," +
                                                "\r\n le nom de la config doit être composée exclusivement de chiffres");
                            return null;
                        }

                        List<String> ListeConfig = new List<String>() { CpRacine.eNomConfiguration() };

                        DicConfig.Add(CpRacine.eKeySansConfig(), ListeConfig);

                        if (ReinitialiserNoDossier)
                        {
                            var nomConfigBase = MdlBase.eNomConfigActive();
                            foreach (var cfg in ListeConfig)
                            {
                                MdlBase.ShowConfiguration2(cfg);
                                MdlBase.eComposantRacine().eEffacerNoDossier();
                            }
                            MdlBase.ShowConfiguration2(nomConfigBase);
                        }
                    }

                    eTypeCorps Filtre = PrendreEnCompteTole ? eTypeCorps.Barre | eTypeCorps.Tole : eTypeCorps.Barre;

                    // Si c'est un assemblage, on liste les composants
                    if (MdlBase.TypeDoc() == eTypeDoc.Assemblage)
                        ListeCp = MdlBase.eRecListeComposant(
                        c =>
                        {

                            if ((ComposantsExterne || c.eEstDansLeDossier(MdlBase)) && !c.IsHidden(true) && !c.ExcludeFromBOM && (c.TypeDoc() == eTypeDoc.Piece))
                            {

                                if (!c.eNomConfiguration().eEstConfigPliee() || DicQte.Plus(c.eKeyAvecConfig()))
                                    return false;

                                if (ReinitialiserNoDossier)
                                    c.eEffacerNoDossier();

                                var LstDossier = c.eListeDesDossiersDePiecesSoudees();
                                foreach (var dossier in LstDossier)
                                {
                                    if (Filtre.HasFlag(dossier.eTypeDeDossier()) && !dossier.eEstExclu())
                                    {
                                        String Materiau = dossier.eGetMateriau();

                                        if (!HashMateriaux.Contains(Materiau))
                                            continue;

                                        if (DicConfig.ContainsKey(c.eKeySansConfig()))
                                        {
                                            List<String> l = DicConfig[c.eKeySansConfig()];
                                            if (l.Contains(c.eNomConfiguration()))
                                                return false;
                                            else
                                            {
                                                l.Add(c.eNomConfiguration());
                                                return true;
                                            }
                                        }
                                        else
                                        {
                                            DicConfig.Add(c.eKeySansConfig(), new List<string>() { c.eNomConfiguration() });
                                            return true;
                                        }
                                    }
                                }
                            }

                            return false;
                        },
                        null,
                        true
                    );

                    for (int noCp = 0; noCp < ListeCp.Count; noCp++)
                    {
                        var Comp = ListeCp[noCp];
                        var NomConfigPliee = Comp.eNomConfiguration();

                        ModelDoc2 mdl = Comp.eModelDoc2();
                        mdl.ShowConfiguration2(NomConfigPliee);

                        if (ReinitialiserNoDossier)
                            mdl.ePartDoc().eReinitialiserNoDossierMax();

                        List<Feature> ListeDossier = mdl.ePartDoc().eListeDesFonctionsDePiecesSoudees(null);

                        for (int noD = 0; noD < ListeDossier.Count; noD++)
                        {
                            Feature f = ListeDossier[noD];
                            BodyFolder dossier = f.GetSpecificFeature2();

                            if (dossier.IsNull() || (dossier.GetBodyCount() == 0) || dossier.eEstExclu()) continue;

                            String Profil = dossier.eProp(CONSTANTES.PROFIL_NOM);

                            if (String.IsNullOrWhiteSpace(Profil)) continue;

                            Body2 Barre = dossier.ePremierCorps();
                            String Materiau = Barre.eGetMateriauCorpsOuComp(Comp);

                            if(Comp.IsRoot())
                                Materiau = Barre.eGetMateriauCorpsOuPiece(mdl.ePartDoc(), NomConfigPliee);

                            if (!HashMateriaux.Contains(Materiau)) continue;

                            Materiau = ForcerMateriau.IsRefAndNotEmpty(Materiau);

                            listeProfil.AjouterProfil(Materiau, Profil);
                        }
                    }

                    ExecuterEn();
                    return listeProfil;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }

                return null;
            }

            protected override void Command()
            {
                try
                {
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            public class ListeProfil
            {
                private Dictionary<String, Dictionary<String, Boolean>> _dic = new Dictionary<string, Dictionary<string, Boolean>>();

                public Dictionary<String, Dictionary<String, Boolean>> DicProfil { get { return _dic; } }

                public ListeProfil() { }

                public void AjouterProfil(String materiau, String profil)
                {
                    var DicMat = new Dictionary<String, Boolean>();

                    if (_dic.ContainsKey(materiau))
                        DicMat = _dic[materiau];
                    else
                        _dic.Add(materiau, DicMat);

                    if (!DicMat.ContainsKey(profil))
                    {
                        DicMat.Add(profil, true);
                        NbProfils++;
                    }
                }

                public int NbProfils { get; private set; }
            }
        }
    }
}


