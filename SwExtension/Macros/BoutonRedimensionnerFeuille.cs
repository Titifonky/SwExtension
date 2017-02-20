using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Redimensionner la feuille"),
        ModuleNom("RedimensionnerFeuille")]
    public class BoutonRedimensionnerFeuille : BoutonBase
    {
        protected override void Command()
        {
            DrawingDoc dessin = App.ModelDoc2.eDrawingDoc();

            dessin.eParcourirLesFeuilles(
                f =>
                {
                    dessin.ActivateSheet(f.GetName());
                    WindowLog.Ecrire("  - " + f.GetName());
                    f.eAjusterAutourDesVues();
                    App.ModelDoc2.eZoomEtendu();
                    return false;
                }
                );
        }
    }
}
