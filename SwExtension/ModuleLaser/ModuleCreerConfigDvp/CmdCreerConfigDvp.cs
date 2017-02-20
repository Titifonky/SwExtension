using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ModuleLaser.ModuleCreerConfigDvp
{
    public class CmdCreerConfigDvp : Cmd
    {
        private ModelDoc2 _mdlBase = null;
        private String nomConfigBase = "";

        public ModelDoc2 MdlBase
        {
            get { return _mdlBase; }
            set { _mdlBase = value; nomConfigBase = value.eNomConfigActive(); }
        }

        public Boolean SupprimerLesAnciennesConfigs = false;
        public Boolean ReconstuireLesConfigs = false;

        private Boolean _ToutesLesConfigurations = false;
        public Boolean ToutesLesConfigurations
        {
            get { return _ToutesLesConfigurations; }
            set
            {
                if (ReinitialiserNoDossier)
                    _ToutesLesConfigurations = true;
                else
                    _ToutesLesConfigurations = value;
            }
        }
        public Boolean MasquerEsquisses = false;
        public Boolean NumeroterDossier = false;

        private Boolean _ReinitialiserNoDossier = false;
        public Boolean ReinitialiserNoDossier
        {
            get { return _ReinitialiserNoDossier; }
            set
            {
                _ReinitialiserNoDossier = value;

                if (value)
                    _ToutesLesConfigurations = true;
            }
        }
        public Boolean MajListePiecesSoudees = false;
        public Boolean SupprimerFonctions = false;
        public String NomFonctionSupprimer = "";

        private List<String> DicErreur = new List<String>();

        protected override void Command()
        {
            try
            {
                List<Component2> ListeCp = new List<Component2>();
                Dictionary<String, List<String>> DicConfig = new Dictionary<String, List<String>>();
                int NbDvp = 0;

                if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                {
                    ListeCp.Add(MdlBase.eComposantRacine());

                    List<String> ListeConfig = new List<String>() { MdlBase.eNomConfigActive().eEstConfigPliee() ? MdlBase.eNomConfigActive() : "" };
                    if (ToutesLesConfigurations)
                        ListeConfig = MdlBase.eListeNomConfiguration(eTypeConfig.Pliee);

                    if (ListeConfig.Count == 0)
                    {
                        WindowLog.Ecrire("Pas de configuration de tole pliee");
                        return;
                    }

                    if (ReinitialiserNoDossier)
                    {
                        foreach (var cfg in ListeConfig)
                        {
                            MdlBase.ShowConfiguration2(cfg);
                            MdlBase.eComposantRacine().eEffacerNoDossier();
                        }
                        MdlBase.ShowConfiguration2(nomConfigBase);
                    }

                    DicConfig.Add(MdlBase.eComposantRacine().eKeySansConfig(), ListeConfig);
                }

                // Si c'est un assemblage, on liste les composants
                if (MdlBase.TypeDoc() == eTypeDoc.Assemblage)
                    ListeCp = MdlBase.eRecListeComposant(
                        c =>
                        {
                            if (!c.IsHidden(true) && (c.TypeDoc() == eTypeDoc.Piece))
                            {
                                if (!c.eNomConfiguration().eEstConfigPliee())
                                    return false;

                                if (DicConfig.ContainsKey(c.eKeySansConfig()))
                                {
                                    if (DicConfig[c.eKeySansConfig()].AddIfNotExist(c.eNomConfiguration()) && ReinitialiserNoDossier)
                                        c.eEffacerNoDossier();

                                    return false;
                                }

                                if (ReinitialiserNoDossier)
                                    c.eEffacerNoDossier();

                                DicConfig.Add(c.eKeySansConfig(), new List<string>() { c.eNomConfiguration() });
                                return true;
                            }

                            return false;
                        },
                        null,
                        true);

                for (int noCp = 0; noCp < ListeCp.Count; noCp++)
                {
                    var Cp = ListeCp[noCp];
                    WindowLog.SautDeLigne();
                    WindowLog.EcrireF("[{2}/{3}] {0}{1}", Cp.eNomSansExt(), !ToutesLesConfigurations ? " \"" + Cp.eNomConfiguration() + "\" " : "", noCp + 1, ListeCp.Count);

                    ModelDoc2 mdl = Cp.eModelDoc2();
                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);

                    List<String> ListeNomConfigs = DicConfig[Cp.eKeySansConfig()];
                    if (ToutesLesConfigurations)
                        ListeNomConfigs = Cp.eModelDoc2().eListeNomConfiguration((String c) => { return c.eEstConfigPliee(); });

                    ListeNomConfigs.Sort(new WindowsStringComparer());

                    if (MasquerEsquisses)
                        cmdMasquerEsquisses(mdl);

                    if (SupprimerLesAnciennesConfigs)
                        cmdSupprimerLesAnciennesConfigs(mdl);

                    if (!mdl.Extension.LinkedDisplayState)
                    {
                        mdl.Extension.LinkedDisplayState = true;

                        foreach (var c in mdl.eListeConfigs(eTypeConfig.Tous))
                            c.eRenommerEtatAffichage();
                    }

                    for (int noCfg = 0; noCfg < ListeNomConfigs.Count; noCfg++)
                    {
                        var NomConfigPliee = ListeNomConfigs[noCfg];
                        WindowLog.SautDeLigne();
                        WindowLog.EcrireF("  [{1}/{2}] Config : \"{0}\"", NomConfigPliee, noCfg + 1, ListeNomConfigs.Count);
                        mdl.ShowConfiguration2(NomConfigPliee);
                        mdl.EditRebuild3();

                        if (SupprimerFonctions)
                            cmdSupprimerFonctions(mdl, NomConfigPliee);

                        PartDoc Piece = mdl.ePartDoc();

                        if (NumeroterDossier)
                            Piece.eNumeroterDossier(MajListePiecesSoudees);

                        var ListeDossier = Piece.eListePIDdesFonctionsDePiecesSoudees(null);

                        for (int noD = 0; noD < ListeDossier.Count; noD++)
                        {
                            var f = ListeDossier[noD];
                            BodyFolder dossier = f.GetSpecificFeature2();

                            if (dossier.IsNull() || (dossier.GetBodyCount() == 0)) continue;

                            WindowLog.EcrireF("    - [{1}/{2}] Dossier : \"{0}\"", f.Name, noD + 1, ListeDossier.Count);

                            Body2 Tole = dossier.eCorpsDeTolerie();

                            if (Tole.IsNull())
                                continue;

                            var pidTole = new SwObjectPID<Body2>(Tole, MdlBase);

                            String NoDossier = dossier.eProp(CONSTANTES.NO_DOSSIER);

                            if (NoDossier.IsNull() || String.IsNullOrWhiteSpace(NoDossier))
                                NoDossier = Piece.eNumeroterDossier(MajListePiecesSoudees)[dossier.eNomDossier()].ToString();

                            String NomConfigDepliee = Sw.eNomConfigDepliee(NomConfigPliee, NoDossier);

                            WindowLog.EcrireF("        cfg : \"{0}\"", NomConfigDepliee);

                            Configuration CfgDepliee = null;

                            if (ReconstuireLesConfigs)
                                CfgDepliee = MdlBase.GetConfigurationByName(NomConfigDepliee);

                            if (CfgDepliee.IsNull())
                                CfgDepliee = mdl.CreerConfigDepliee(NomConfigDepliee, NomConfigPliee);
                            else if (!ReconstuireLesConfigs)
                                continue;

                            if (CfgDepliee.IsNull())
                            {
                                WindowLog.Ecrire("       - Config non crée");
                                continue;
                            }

                            NbDvp++;

                            try
                            {
                                mdl.ShowConfiguration2(NomConfigDepliee);

                                // On ajoute le numero de la config parent aux propriétés de la configuration depliée
                                CfgDepliee.CustomPropertyManager.Add3(CONSTANTES.NO_CONFIG, (int)swCustomInfoType_e.swCustomInfoText, NomConfigPliee, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);

                                pidTole.Maj(ref Tole);

                                Tole.DeplierTole(mdl, NomConfigDepliee);
                            }
                            catch (Exception e)
                            {
                                DicErreur.Add(Cp.eNomSansExt() + " -> cfg : " + NomConfigPliee + " - No : " + NoDossier + " = " + NomConfigDepliee);
                                WindowLog.Ecrire("Erreur de depliage");
                                this.LogMethode(new Object[] { e });
                            }

                            try
                            {
                                mdl.ShowConfiguration2(NomConfigPliee);

                                pidTole.Maj(ref Tole);

                                Tole.PlierTole(mdl, NomConfigPliee);
                            }
                            catch (Exception e)
                            {
                                WindowLog.Ecrire("Erreur de repliage");
                                this.LogMethode(new Object[] { e });
                            }
                        }
                    }

                    Cp.eModelDoc2().ShowConfiguration2(Cp.eNomConfiguration());
                    WindowLog.SautDeLigne();

                    if (Cp.GetPathName() != MdlBase.GetPathName())
                        App.Sw.CloseDoc(mdl.GetPathName());
                }

                if (DicErreur.Count > 0)
                {
                    WindowLog.SautDeLigne();

                    WindowLog.Ecrire("Liste des erreurs :");
                    foreach (var item in DicErreur)
                        WindowLog.Ecrire(" - " + item);

                    WindowLog.SautDeLigne();
                }
                else
                {
                    WindowLog.Ecrire("Pas d'erreur");
                }

                WindowLog.SautDeLigne();
                WindowLog.Ecrire("Resultat :");
                WindowLog.Ecrire("----------------");
                WindowLog.EcrireF("  {0} dvp crées", NbDvp);

                MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                MdlBase.ShowConfiguration2(nomConfigBase);
                MdlBase.EditRebuild3();
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private void cmdMasquerEsquisses(ModelDoc2 mdl)
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

        private void cmdSupprimerLesAnciennesConfigs(ModelDoc2 mdl)
        {
            WindowLog.SautDeLigne();

            if (mdl.eNomConfigActive().eEstConfigDepliee())
                mdl.ShowConfiguration2(mdl.eListeNomConfiguration()[0]);

            mdl.EditRebuild3();

            WindowLog.Ecrire("  - Suppression des cfgs depliées :");
            var liste = mdl.eListeConfigs(eTypeConfig.Depliee);
            if (liste.Count == 0)
                WindowLog.EcrireF("   Aucune configuration à supprimer");

            foreach (Configuration Cf in liste)
            {
                String IsSup = Cf.eSupprimerConfigAvecEtatAff(mdl) ? "Ok" : "Erreur";
                WindowLog.EcrireF("  {0} : {1}", Cf.Name, IsSup);
            }
        }

        private void cmdSupprimerFonctions(ModelDoc2 mdl, String nomConfigDepliee)
        {
            if (String.IsNullOrWhiteSpace(NomFonctionSupprimer))
                return;

            mdl.eParcourirFonctions(
                f =>
                {
                    if (Regex.IsMatch(f.Name, NomFonctionSupprimer))
                        f.eModifierEtat(swFeatureSuppressionAction_e.swSuppressFeature, nomConfigDepliee);

                    return false;
                },
                false
                );
        }

    }
}


