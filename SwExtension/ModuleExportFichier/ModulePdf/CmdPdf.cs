using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.IO;

namespace ModuleExportFichier
{
    namespace ModulePdf
    {
        public class CmdPdf : Cmd
        {
            public DrawingDoc Dessin = null;

            public Boolean ToutesLesFeuilles = false;
            public eTypeFichierExport typeExport;
            public String CheminDossier;
            public String NomFichier;
            public Sheet Feuille;
            public String CheminFichierExport;

            protected override void Command()
            {
                try
                {
                    WindowLog.EcrireF("Export : {0}", typeExport.GetEnumInfo<Intitule>().ToUpperInvariant());
                    WindowLog.SautDeLigne();
                    WindowLog.EcrireF("  Dossier : {0}", new DirectoryInfo(CheminDossier).Name);

                    WindowLog.EcrireF("   {0}", NomFichier + typeExport.GetEnumInfo<ExtFichier>());
                    CheminFichierExport = Feuille.eExporterEn(Dessin, typeExport, CheminDossier, NomFichier, ToutesLesFeuilles);
                }
                catch (Exception e)
                {
                    WindowLog.Ecrire("Erreur : Consultez le fichier LOG");
                    this.LogMethode(new Object[] { e });
                }
            }
        }
    }
}


