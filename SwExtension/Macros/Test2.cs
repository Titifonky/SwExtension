using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Test2"),
        ModuleNom("Test2")]

    public class Test2 : BoutonBase
    {
        protected override void Command()
        {
            try
            {
                ModelDoc2 mdl = App.ModelDoc2;

                
            }
            catch (Exception e) { this.LogMethode(new Object[] { e }); }

        }

        //protected override void Command()
        //{
        //    try
        //    {
        //        ModelDoc2 mdl = App.ModelDoc2;

        //        var Face = mdl.eSelect_RecupererObjet<Face2>(1);
        //        Body2 Corps = Face.GetBody();

        //        mdl.eEffacerSelection();

        //        List<Face2> ListeFaceExt = new List<Face2>();

        //        foreach (var f in Corps.eListeDesFaces())
        //        {
        //            Byte[] Tab = mdl.Extension.GetPersistReference3(f);
        //            String S = System.Text.Encoding.Default.GetString(Tab);

        //            int Pos_moSideFace = S.IndexOf("moSideFace3IntSurfIdRep_c");

        //            int Pos_moVertexRef = S.Position("moVertexRef");

        //            int Pos_moDerivedSurfIdRep = S.Position("moDerivedSurfIdRep_c");

        //            int Pos_moFromSkt = Math.Min(S.Position("moFromSktEntSurfIdRep_c"), S.Position("moFromSktEnt3IntSurfIdRep_c"));

        //            int Pos_moEndFace = Math.Min(S.Position("moEndFaceSurfIdRep_c"), S.Position("moEndFace3IntSurfIdRep_c"));

        //            if (Pos_moSideFace != -1 && Pos_moSideFace < Pos_moEndFace && Pos_moSideFace < Pos_moFromSkt && Pos_moSideFace < Pos_moVertexRef && Pos_moSideFace < Pos_moDerivedSurfIdRep)
        //                ListeFaceExt.Add(f);

        //            Log.Message(S);
        //            Log.MessageF("Side {0} From {1} End {2}", Pos_moSideFace, Pos_moFromSkt, Pos_moEndFace);
        //        }

        //        foreach (var f in ListeFaceExt)
        //            f.eSelectEntite(true);
        //    }
        //    catch (Exception e) { this.LogMethode(new Object[] { e }); }

        //}

        //protected override void Command()
        //{
        //    try
        //    {
        //        ModelDoc2 mdl = App.ModelDoc2;
        //        //var DossierExport = mdl.eDossier();
        //        //var NomFichier = mdl.eNomSansExt();

        //        //var F = mdl.eSelect_RecupererObjet<Feature>(1);
        //        //mdl.eEffacerSelection();

        //        //var def = (StructuralMemberFeatureData)F.GetDefinition();
        //        //WindowLog.Ecrire(def.WeldmentProfilePath);
        //        //WindowLog.Ecrire(def.ConfigurationName);
        //        //foreach (var sf in F.eListeSousFonction())
        //        //{
        //        //    WindowLog.Ecrire(sf.GetTypeName2());
        //        //}

        //        //var Face = mdl.eSelect_RecupererObjet<Face2>(1);

        //        //Byte[] Tab = mdl.Extension.GetPersistReference3(Face);
        //        //String S = System.Text.Encoding.Default.GetString(Tab);
        //        //Log.Message(S);

        //        var Face = mdl.eSelect_RecupererObjet<Face2>(1);
        //        Body2 Corps = Face.GetBody();

        //        mdl.eEffacerSelection();

        //        List<Face2> ListeFaceExt = new List<Face2>();

        //        foreach (var f in Corps.eListeDesFaces())
        //        {
        //            Byte[] Tab = mdl.Extension.GetPersistReference3(f);
        //            String S = System.Text.Encoding.Default.GetString(Tab);

        //            int Pos_moSideFace = S.IndexOf("moSideFace3IntSurfIdRep_c");

        //            int Pos_moFromSkt = Math.Min(S.Position("moFromSktEntSurfIdRep_c"), S.Position("moFromSktEnt3IntSurfIdRep_c"));

        //            int Pos_moEndFace = Math.Min(S.Position("moEndFaceSurfIdRep_c"), S.Position("moEndFace3IntSurfIdRep_c"));

        //            Log.Message(S);
        //            Log.MessageF("Side {0} From {1} End {2}", Pos_moSideFace, Pos_moFromSkt, Pos_moEndFace);

        //            if (Pos_moSideFace != -1 && Pos_moSideFace < Pos_moEndFace && Pos_moSideFace < Pos_moFromSkt)
        //                ListeFaceExt.Add(f);
        //        }

        //        foreach (var f in ListeFaceExt)
        //            f.eSelectEntite(true);
        //    }
        //    catch (Exception e) { this.LogMethode(new Object[] { e }); }

        //}

    }
}
