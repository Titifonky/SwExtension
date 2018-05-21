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
                var DicCorpsTest = new Dictionary<Body2, Tuple<String, String>>();
                var DicDossier = new Dictionary<Body2, Tuple<String, int>>();

                foreach (var comp in mdl.eComposantRacine().eRecListeComposant())
                {
                    var ListefDossier = comp.eListeDesFonctionsDePiecesSoudees();

                    for (int i = 0; i < ListefDossier.Count; i++)
                    {
                        Feature fDossier = ListefDossier[i];
                        BodyFolder Dossier = fDossier.GetSpecificFeature2();
                        if (Dossier.IsNull() || Dossier.eNbCorps() == 0 || Dossier.eEstExclu() || !Dossier.eEstUnDossierDeBarres())
                            continue;

                        var Corps = Dossier.ePremierCorps();
                        if (Corps.IsNull()) continue;

                        var MateriauCorps = Corps.eGetMateriauCorpsOuComp(comp);

                        Boolean Ajoute = false;
                        foreach (var CorpsTest in DicCorpsTest.Keys)
                        {
                            var MateriauCorpsTest = DicCorpsTest[CorpsTest].Item1;
                            if (MateriauCorps != MateriauCorpsTest) continue;

                            if (Corps.eEstSemblable(CorpsTest))
                            {
                                var t = DicDossier[CorpsTest];

                                var nb = t.Item2 + Dossier.eNbCorps();

                                DicDossier[CorpsTest] = new Tuple<String, int>(t.Item1, nb);

                                Ajoute = true;
                                break;
                            }

                        }

                        if (Ajoute == false)
                        {
                            var FonctionPID = new SwObjectPID<Feature>(fDossier, comp.eModelDoc2());
                            DicCorpsTest.Add(Corps, new Tuple<string, String>(MateriauCorps, fDossier.Name));
                            DicDossier.Add(Corps, new Tuple<String, int>(fDossier.Name, Dossier.eNbCorps()));
                        }
                    }

                }

                WindowLog.EcrireF("Nb de corps unique : {0}", DicDossier.Count);
                int nbtt = 0;
                foreach (var t in DicDossier.Values)
                {
                    nbtt += t.Item2;
                    WindowLog.EcrireF("{0} : {1}", t.Item1, t.Item2);
                }

                WindowLog.EcrireF("Nb total de corps : {0}", nbtt);
            }
            catch (Exception e) { this.LogMethode(new Object[] { e }); }

        }

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
