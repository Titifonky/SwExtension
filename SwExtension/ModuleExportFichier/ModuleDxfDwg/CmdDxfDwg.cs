using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.IO;

namespace ModuleExportFichier
{
    namespace ModuleDxfDwg
    {
        public class CmdDxfDwg : Cmd
        {
            public DrawingDoc Dessin = null;

            public eTypeFichierExport typeExport;
            public String CheminDossier;
            public String NomFichier;
            public Sheet Feuille;
            public Boolean ToutesLesFeuilles = false;

            protected override void Command()
            {
                try
                {
                    WindowLog.EcrireF("Export : {0}", typeExport.GetEnumInfo<Intitule>().ToUpperInvariant());
                    WindowLog.SautDeLigne();
                    WindowLog.EcrireF("  Dossier : {0}", new DirectoryInfo(CheminDossier).Name);

                    WindowLog.EcrireF("   {0}", NomFichier + typeExport.GetEnumInfo<ExtFichier>());
                    Feuille.eExporterEn(Dessin, typeExport, CheminDossier, NomFichier, ToutesLesFeuilles);
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


