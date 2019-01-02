using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace ModuleProduction.ModuleGenererConfigDvp
{
    public class CmdGenererConfigDvp : Cmd
    {
        public ModelDoc2 MdlBase
        {
            get { return _mdlBase; }
            set { _mdlBase = value; }
        }

        private ModelDoc2 _mdlBase = null;

        public Boolean SupprimerLesAnciennesConfigs = false;
        public Boolean MasquerEsquisses = false;

        protected override void Command()
        {
            try
            {
                var ListeCorps = MdlBase.pChargerNomenclature();

                foreach (var corps in ListeCorps.Values)
                {
                    WindowLog.EcrireF("{0} -> dvp", corps.Repere);
                    CreerDvp(corps, MdlBase.pDossierPiece(), SupprimerLesAnciennesConfigs, MasquerEsquisses);
                }

                MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                MdlBase.EditRebuild3();
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        public static void CreerDvp(Corps corps, String dossierPiece, Boolean _supprimerLesAnciennesConfigs = false, Boolean _masquerEsquisses = false)
        {
            try
            {
                if (corps.TypeCorps != eTypeCorps.Tole) return;

                String Repere = CONSTANTES.PREFIXE_REF_DOSSIER + corps.Repere;

                var nomFichier = Repere + OutilsProd.ExtPiece;
                var chemin = Path.Combine(dossierPiece, nomFichier);
                if (!File.Exists(chemin)) return;
                
                var mdl = Sw.eOuvrir(chemin);
                if (mdl.IsNull()) return;

                var NomCfgPliee = mdl.eNomConfigActive();
                var Piece = mdl.ePartDoc();
                var Tole = Piece.ePremierCorps();

                if (_supprimerLesAnciennesConfigs)
                    cmdSupprimerLesAnciennesConfigs(mdl);

                if (_masquerEsquisses)
                    cmdMasquerEsquisses(mdl);

                if (!mdl.Extension.LinkedDisplayState)
                {
                    mdl.Extension.LinkedDisplayState = true;

                    foreach (var c in mdl.eListeConfigs(eTypeConfig.Tous))
                        c.eRenommerEtatAffichage();
                }

                //mdl.UnlockAllExternalReferences();
                mdl.EditRebuild3();

                String NomConfigDepliee = Sw.eNomConfigDepliee(NomCfgPliee, Repere);

                if (!mdl.pCreerConfigDepliee(NomConfigDepliee, NomCfgPliee))
                {
                    WindowLog.Ecrire("       - Config non crée");
                    return;
                }
                try
                {
                    mdl.ShowConfiguration2(NomConfigDepliee);
                    mdl.EditRebuild3();
                    Piece.pDeplierTole(NomConfigDepliee);

                    mdl.ShowConfiguration2(NomCfgPliee);
                    mdl.EditRebuild3();
                    Piece.pPlierTole(NomCfgPliee);
                    WindowLog.EcrireF("  - Dvp crée : {0}", NomConfigDepliee);
                }
                catch (Exception e)
                {
                    WindowLog.Ecrire("  - Erreur de dvp");
                    Log.Message(new Object[] { e });
                }

                mdl.ShowConfiguration2(NomCfgPliee);
                mdl.EditRebuild3();
                mdl.eSauver();
            }
            catch (Exception e)
            {
                Log.Message(new Object[] { e });
            }
        }

        private static void cmdMasquerEsquisses(ModelDoc2 mdl)
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

        private static void cmdSupprimerLesAnciennesConfigs(ModelDoc2 mdl)
        {
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
    }
}


