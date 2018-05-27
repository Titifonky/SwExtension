using LogDebugging;
using Outils;
using SwExtension;
using System;
using System.IO;
using System.Reflection;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Afficher le log de debug"),
        ModuleNom("AfficherLogDebug")]
    public class BoutonAfficherLogDebug : BoutonBase
    {
        private Parametre NomFichierLogDebug;

        public BoutonAfficherLogDebug()
        {
            LogToWindowLog = false;

            NomFichierLogDebug = _Config.AjouterParam("Fichier", "LOGs.log");
        }

        protected override void Command()
        {
            var dossier = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var cheminFichier = Path.Combine(dossier, NomFichierLogDebug.GetValeur<String>());

            if (File.Exists(cheminFichier))
                System.Diagnostics.Process.Start(cheminFichier);
        }
    }
}
