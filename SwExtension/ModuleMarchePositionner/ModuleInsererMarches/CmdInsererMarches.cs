using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace ModuleMarchePositionner
{
    namespace ModuleInsererMarches
    {
        public class CmdInsererMarches : Cmd
        {
            public ModelDoc2 MdlBase = null;
            public Component2 Marche = null;
            public Component2 ComposantRepetition = null;
            public Feature FonctionRepetition = null;
            public Entity Point = null;
            public Entity Arrete = null;
            public Feature Plan = null;
            public SketchPoint PointMarche = null;
            public Feature AxeMarche = null;
            public Feature PlanMarche = null;

            private AssemblyDoc Ass = null;

            protected override void Command()
            {
                Ass = MdlBase.eAssemblyDoc();

                try
                {
                    CurveDrivenPatternFeatureData def = FonctionRepetition.GetDefinition();
                    Object[] tabB = def.PatternBodyArray;
                    Body2 b = (Body2)tabB[0];

                    MathTransform compRepetTrans = ComposantRepetition.Transform2;
                    MathTransform baseMarcheTrans = Marche.Transform2;

                    Vertex v = (Vertex)Point;
                    Edge e = (Edge)Arrete;

                    List<Face2> ListeFace = FonctionRepetition.eListeDesFaces();

                    Sketch sk = PointMarche.GetSketch();
                    Feature fPointMarche = (Feature)sk;

                    List<Component2> ListeComposants = new List<Component2>() { Marche };

                    Double arr = 0.0001;

                    for (int i = 1; i <= (def.D1InstanceCount + def.D2InstanceCount); i++)
                    {
                        MathTransform bodyTrans = def.GetTransform(i);
                        if (bodyTrans.IsNull()) break;

                        MathTransform Transform = compRepetTrans.Inverse();
                        Transform = Transform.Multiply(bodyTrans);
                        Transform = Transform.Multiply(compRepetTrans);
                        Transform = baseMarcheTrans.Multiply(Transform);

                        Point pt = new Point(v);
                        Segment sg = new Segment(e);
                        pt.MultiplyTransfom(bodyTrans);
                        sg.MultiplyTransfom(bodyTrans);

                        Vertex vertex = null;
                        Edge edge = null;

                        foreach (Face2 face in ListeFace)
                        {
                            foreach (Edge ed in face.eListeDesArretes())
                            {
                                Segment s = new Segment(ed);

                                if (sg.Compare(s, arr))
                                    edge = ed;

                                if (s.Start.Comparer(pt, arr))
                                    vertex = ed.GetStartVertex();

                                if (s.End.Comparer(pt, arr))
                                    vertex = ed.GetEndVertex();

                                if (edge.IsRef() && vertex.IsRef())
                                    break;
                            }

                            if (edge.IsRef() && vertex.IsRef())
                                break;
                        }

                        Entity eVertex = (Entity)vertex;
                        Entity eEdge = (Entity)edge;

                        Component2 cp = Ass.AddComponent5(Marche.GetPathName(), (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);
                        if (cp.IsNull()) continue;

                        // Quand on réinsere un composant, les précédentes contraintes sont recrées.
                        // On les supprime pour eviter des conflits
                        Object[] Mates = cp.GetMates();

                        if (Mates.IsRef())
                        {
                            foreach (Feature mate in Mates)
                            {
                                mate.eSelect();
                                MdlBase.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                            }
                        }

                        WindowLog.Ecrire(cp.Name2);

                        cp.Transform2 = Transform;
                        ListeComposants.Add(cp);

                        int longstatus = 0;

                        MdlBase.eEffacerSelection();
                        eVertex.eSelectEntite(false);
                        Feature f = cp.FeatureByName(fPointMarche.Name);
                        Sketch sketchOrigine = f.GetSpecificFeature2();
                        Object[] tabPt = sketchOrigine.GetSketchPoints2();
                        SketchPoint origine = (SketchPoint)tabPt[0];
                        origine.Select4(true, null);
                        Mate2 mPoint = Ass.AddMate5((int)swMateType_e.swMateCOINCIDENT,
                                     (int)swMateAlign_e.swMateAlignCLOSEST, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out longstatus);
                        MdlBase.eEffacerSelection();

                        MdlBase.eEffacerSelection();
                        eEdge.eSelectEntite(false);
                        cp.FeatureByName(AxeMarche.Name).eSelect(true);
                        Mate2 mAxe = Ass.AddMate5((int)swMateType_e.swMateANGLE,
                                     (int)swMateAlign_e.swAlignNONE, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out longstatus);
                        MdlBase.eEffacerSelection();

                        MdlBase.eEffacerSelection();
                        Plan.eSelect(false);
                        cp.FeatureByName(PlanMarche.Name).eSelect(true);
                        Mate2 mPlan = Ass.AddMate5((int)swMateType_e.swMatePARALLEL,
                                     (int)swMateAlign_e.swAlignNONE, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out longstatus);
                        MdlBase.eEffacerSelection();

                        Feature m = (Feature)mPoint;
                        WindowLog.Ecrire("   " + m.Name);
                        m = (Feature)mAxe;
                        WindowLog.Ecrire("   " + m.Name);
                        m = (Feature)mPlan;
                        WindowLog.Ecrire("   " + m.Name);

                    }

                    MdlBase.eEffacerSelection();

                    foreach (var cp in ListeComposants)
                    {
                        cp.eSelectById(MdlBase, -1, true);
                    }

                    Feature Dossier = MdlBase.FeatureManager.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing);
                    Dossier.eRenommerFonction(String.Format("Marches ({0} {1}) ", ComposantRepetition.Name2, FonctionRepetition.Name));
                }
                catch (Exception e)
                {
                    this.LogMethode(new Object[] { e });
                }
            }
        }
    }
}


