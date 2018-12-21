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
            set { _mdlBase = value; nomConfigBase = value.eNomConfigActive(); }
        }

        private ModelDoc2 _mdlBase = null;
        private String nomConfigBase = "";

        public Boolean SupprimerLesAnciennesConfigs = false;
        public Boolean MasquerEsquisses = false;

        private List<String> DicErreur = new List<String>();

        protected override void Command()
        {
            try
            {
                var ListeCorps = MdlBase.ChargerNomenclature();

                foreach (var corps in ListeCorps.Values)
                {
                    if (corps.TypeCorps != eTypeCorps.Tole) continue;

                    String Repere = CONSTANTES.PREFIXE_REF_DOSSIER + corps.Repere;

                    var chemin = Path.Combine(MdlBase.DossierPiece(), Repere + OutilsProd.ExtPiece);
                    if (!File.Exists(chemin)) continue;

                    var mdl = Sw.eOuvrir(chemin);
                    if (mdl.IsNull()) continue;

                    WindowLog.EcrireF("{0}", Repere);

                    var NomCfgPliee = mdl.eNomConfigActive();
                    var Piece = mdl.ePartDoc();
                    var Tole = Piece.ePremierCorps();

                    if (SupprimerLesAnciennesConfigs)
                        cmdSupprimerLesAnciennesConfigs(mdl);

                    if (MasquerEsquisses)
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

                    if (!mdl.CreerConfigDepliee(NomConfigDepliee, NomCfgPliee))
                    {
                        WindowLog.Ecrire("       - Config non crée");
                        continue;
                    }
                    try
                    {
                        mdl.ShowConfiguration2(NomConfigDepliee);
                        mdl.EditRebuild3();
                        Piece.DeplierTole(NomConfigDepliee);

                        mdl.ShowConfiguration2(NomCfgPliee);
                        mdl.EditRebuild3();
                        Piece.PlierTole(NomCfgPliee);
                        WindowLog.EcrireF("  - Dvp crée : {0}", NomConfigDepliee);
                    }
                    catch (Exception e)
                    {
                        DicErreur.Add(String.Format("{0} -> Erreur", mdl.eNomSansExt()));
                        WindowLog.Ecrire("  - Erreur de dvp");
                        this.LogMethode(new Object[] { e });
                    }

                    mdl.ShowConfiguration2(NomCfgPliee);
                    //mdl.LockAllExternalReferences();
                    mdl.EditRebuild3();
                    mdl.eSauver();
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
                    WindowLog.SautDeLigne();
                    WindowLog.Ecrire("Pas d'erreur");
                }

                MdlBase.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                MdlBase.ShowConfiguration2(nomConfigBase);
                MdlBase.EditRebuild3();
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
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


