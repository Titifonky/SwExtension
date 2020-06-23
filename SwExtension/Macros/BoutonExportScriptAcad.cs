using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.IO;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Export esquisse script acad"),
        ModuleNom("ExportScriptAcad")]
    public class ExportScriptAcad : BoutonBase
    {
        public ExportScriptAcad() { }

        protected override void Command()
        {
            try
            {
                WindowLog.Ecrire("Export Script");
                var fonction = MdlBase.eSelect_RecupererObjet<Feature>(1);
                var comp = MdlBase.eSelect_RecupererComposant();

                if (fonction.IsNull()) return;
                
                var esquisse = fonction.GetSpecificFeature2() as Sketch;
                if (esquisse.IsNull()) return;

                var nbLignes = esquisse.GetLineCount2((short)swCrossHatchFilter_e.swCrossHatchExclude);
                var arrLignes = (Double[])esquisse.GetLines2((short)swCrossHatchFilter_e.swCrossHatchExclude);

                var nl = System.Environment.NewLine;
                var scriptTexte = "OSNAPCOORD" + nl;
                scriptTexte += "1" + nl + nl + nl;

                var arr = (object[])esquisse.GetSketchSegments();
                foreach (SketchSegment sg in arr)
                {
                    if(!sg.ConstructionGeometry && (sg.GetType() == (int)swSketchSegments_e.swSketchLINE))
                    {
                        var skLine = sg as SketchLine;
                        var sp = skLine.GetStartPoint2() as SketchPoint;
                        var ep = skLine.GetEndPoint2() as SketchPoint;

                        scriptTexte += "_line" + nl;
                        scriptTexte += String.Format("{0:0.000},{1:0.000},{2:0.000}", -1.0 * sp.X, sp.Z, sp.Y) + nl;
                        scriptTexte += String.Format("{0:0.000},{1:0.000},{2:0.000}", -1.0 * ep.X, ep.Z, ep.Y) + nl + nl;
                    }
                }

                WindowLog.Ecrire(nbLignes + " ligne(s) exportée(s)");

                scriptTexte += "'_.zoom _e" + nl;

                var nomMdl = MdlBase.eNomSansExt();

                if (comp.IsRef())
                    nomMdl = comp.eModelDoc2().eNomSansExt();

                var chemin = Path.Combine(MdlBase.eDossier(), nomMdl + "_" + fonction.Name + "_ScriptExportLigne.scr");

                File.WriteAllText(chemin, scriptTexte);

                WindowLog.Ecrire("fichier : " + chemin);
            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }
    }
}
