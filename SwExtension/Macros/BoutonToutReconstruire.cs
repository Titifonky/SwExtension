using LogDebugging;
using Outils;
using SwExtension;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Tout reconstruire"),
        ModuleNom("ToutReconstruire")]
    public class BoutonToutReconstruire : BoutonBase
    {
        public BoutonToutReconstruire()
        {
            LogToWindowLog = false;
        }

        protected override void Command()
        {
            App.ModelDoc2.ForceRebuild3(false);
        }
    }
}
