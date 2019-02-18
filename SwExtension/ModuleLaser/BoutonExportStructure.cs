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
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Exporter les structures"),
        ModuleNom("ExportStructure")]

    public class BoutonExportStructure : BoutonBase
    {
        private Parametre NomDossierExport;

        public BoutonExportStructure()
        {
            LogToWindowLog = false;

            NomDossierExport = _Config.AjouterParam("Dossier", "Export structure");
        }

        private String DossierExport = "";
        private String Indice = "Ind A";

        protected override void Command()
        {
            try
            {
                CreerDossierExport(MdlBase);

                var dic = MdlBase.ListerComposants(false);

                int MdlPct = 0;
                foreach (var mdl in dic.Keys)
                {
                    WindowLog.SautDeLigne();
                    WindowLog.EcrireF("[{1}/{2}] {0}", mdl.eNomSansExt(), ++MdlPct, dic.Count);

                    mdl.eActiver(swRebuildOnActivation_e.swRebuildActiveDoc);
                    var cfgActive = mdl.eNomConfigActive();

                    var ListeNomConfigs = dic[mdl];
                    int CfgPct = 0;
                    foreach (var NomConfigPliee in ListeNomConfigs.Keys)
                    {
                        WindowLog.SautDeLigne();
                        WindowLog.EcrireF("  [{1}/{2}] Config : \"{0}\"", NomConfigPliee, ++CfgPct, ListeNomConfigs.Count);
                        mdl.ShowConfiguration2(NomConfigPliee);
                        mdl.EditRebuild3();
                        mdl.eEffacerSelection();

                        var NbConfig = dic[mdl][NomConfigPliee];
                        PartDoc Piece = mdl.ePartDoc();

                        var ListeDossier = Piece.eListeDesDossiersDePiecesSoudees(
                            dossier =>
                            {
                                if (dossier.eEstExclu() || dossier.IsNull() || (dossier.GetBodyCount() == 0))
                                    return false;

                                String Profil = dossier.eProp(CONSTANTES.PROFIL_NOM);
                                String Longueur = dossier.eProp(CONSTANTES.PROFIL_LONGUEUR);

                                if (String.IsNullOrWhiteSpace(Profil) || String.IsNullOrWhiteSpace(Longueur))
                                    return false;

                                foreach (var Barre in dossier.eListeDesCorps())
                                    Barre.Select2(true, null);

                                return true;
                            }
                            );

                        if (ListeDossier.Count == 0) continue;

                        var NomFichier = String.Format("{0}-{1} (x{2}) - {3}", mdl.eNomSansExt(), NomConfigPliee, NbConfig, Indice);
                        
                        var mdlExport = ExportSelection(Piece, NomFichier, eTypeFichierExport.Piece);

                        if (mdlExport.IsNull())
                        {
                            WindowLog.EcrireF("Erreur : {0}", NomFichier);
                            continue;
                        }

                        WindowLog.EcrireF("     {0}", NomFichier);

                        mdlExport.ViewZoomtofit2();
                        mdlExport.ShowNamedView2("*Isométrique", 7);
                        int lErrors = 0, lWarnings = 0;
                        mdlExport.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref lErrors, ref lWarnings);

                        mdlExport.eFermer();

                        mdl.eFermerSiDifferent(MdlBase);
                    }
                }
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }

        private ModelDoc2 ExportSelection(PartDoc piece, String nomFichier, eTypeFichierExport typeExport)
        {
            int pStatut;
            int pWarning;

            Boolean Resultat = piece.SaveToFile3(Path.Combine(DossierExport, nomFichier + typeExport.GetEnumInfo<ExtFichier>()),
                                                  (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                                  (int)swCutListTransferOptions_e.swCutListTransferOptions_CutListProperties,
                                                  false,
                                                  "",
                                                  out pStatut,
                                                  out pWarning);
            if (Resultat)
                return MdlBase;

            return null;
        }

        private void CreerDossierExport(ModelDoc2 MdlBase)
        {
            String NomBase = NomDossierExport.GetValeur<String>();

            DirectoryInfo D = new DirectoryInfo(MdlBase.eDossier());
            List<String> ListeD = new List<string>();

            foreach (var d in D.GetDirectories())
            {
                if (d.Name.ToUpperInvariant().StartsWith(NomBase.ToUpperInvariant()))
                {
                    ListeD.Add(d.Name);
                }
            }

            ListeD.Sort(new WindowsStringComparer(System.ComponentModel.ListSortDirection.Ascending));

            Indice = OutilsCommun.ChercherIndice(ListeD);

            DossierExport = Path.Combine(MdlBase.eDossier(), NomBase + " - " + Indice);

            if (!Directory.Exists(DossierExport))
                Directory.CreateDirectory(DossierExport);
        }
    }
}
