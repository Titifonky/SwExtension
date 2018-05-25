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
            set { _ToutesLesConfigurations = value; }
        }
        public Boolean MasquerEsquisses = false;

        public Boolean MajListePiecesSoudees = false;
        public Boolean SupprimerFonctions = false;
        public String NomFonctionSupprimer = "";

        private List<String> DicErreur = new List<String>();

        protected override void Command()
        {
            try
            {
                int NbDvp = 0;

                var dic = MdlBase.ListerComposants(false, eTypeCorps.Tole);

                int MdlPct = 0;
                foreach (var mdl in dic.Keys)
                {
                    WindowLog.SautDeLigne();
                    WindowLog.EcrireF("[{1}/{2}] {0}", mdl.eNomSansExt(), ++MdlPct, dic.Count);

                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                    var cfgActive = mdl.eNomConfigActive();

                    var ListeNomConfigs = dic[mdl];

                    if (ToutesLesConfigurations)
                        mdl.eParcourirConfiguration(
                            (String c) =>
                            {
                                if (c.eEstConfigPliee())
                                    ListeNomConfigs.Add(c, 1);

                                return false;
                            }
                            );

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

                    int CfgPct = 0;
                    foreach (var NomConfigPliee in ListeNomConfigs.Keys)
                    {
                        WindowLog.SautDeLigne();
                        WindowLog.EcrireF("  [{1}/{2}] Config : \"{0}\"", NomConfigPliee, ++CfgPct, ListeNomConfigs.Count);
                        mdl.ShowConfiguration2(NomConfigPliee);
                        mdl.EditRebuild3();

                        if (SupprimerFonctions)
                            cmdSupprimerFonctions(mdl, NomConfigPliee);

                        PartDoc Piece = mdl.ePartDoc();

                        var ListeDossier = Piece.eListePIDdesFonctionsDePiecesSoudees(null);

                        for (int noD = 0; noD < ListeDossier.Count; noD++)
                        {
                            var f = ListeDossier[noD];
                            BodyFolder dossier = f.GetSpecificFeature2();

                            var RefDossier = dossier.eProp(CONSTANTES.REF_DOSSIER);

                            if (dossier.eEstExclu() || dossier.IsNull() || (dossier.GetBodyCount() == 0)) continue;
                            

                            Body2 Tole = dossier.eCorpsDeTolerie();

                            if (Tole.IsNull())
                                continue;

                            var pidTole = new SwObjectPID<Body2>(Tole, MdlBase);

                            String NomConfigDepliee = Sw.eNomConfigDepliee(NomConfigPliee, RefDossier);

                            WindowLog.EcrireF("    - [{1}/{2}] Dossier : \"{0}\" -> {3}", f.Name, noD + 1, ListeDossier.Count, NomConfigDepliee);

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
                                DicErreur.Add(mdl.eNomSansExt() + " -> cfg : " + NomConfigPliee + " - No : " + RefDossier + " = " + NomConfigDepliee);
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

                    mdl.ShowConfiguration2(cfgActive);
                    WindowLog.SautDeLigne();

                    if (mdl.GetPathName() != MdlBase.GetPathName())
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


