using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace ModuleMarcheConfig
{
    namespace ModulePositionnerPlatine
    {
        public class CmdPositionnerPlatine : Cmd
        {
            public ModelDoc2 MdlBase = null;
            public Face2 F_Dessus = null;
            public Face2 F_Devant = null;
            public Feature PltG_Plan = null;
            public Feature PltD_Plan = null;
            public Component2 PltG = null;
            public Component2 PltD = null;
            public Boolean SurTouteLesConfigs = false;


            private String _NomConfigCourante = "";

            protected override void Command()
            {
                _NomConfigCourante = MdlBase.eNomConfigActive();

                if (SurTouteLesConfigs)
                {
                    var pidF_Dessus = new SwObjectPID<Face2>(F_Dessus, MdlBase);
                    var pidF_Devant = new SwObjectPID<Face2>(F_Devant, MdlBase);
                    var pidPltG_Plan = new SwObjectPID<Feature>(PltG_Plan, MdlBase);
                    var pidPltD_Plan = new SwObjectPID<Feature>(PltD_Plan, MdlBase);
                    var pidPltG = new SwObjectPID<Component2>(PltG, MdlBase);
                    var pidPltD = new SwObjectPID<Component2>(PltD, MdlBase);

                    List<String> ListeNomsConfig = MdlBase.eListeNomConfiguration(eTypeConfig.DeBase);

                    foreach (String NomConfig in ListeNomsConfig)
                    {
                        MdlBase.ShowConfiguration2(NomConfig);
                        MdlBase.EditRebuild3();

                        //pidF_Dessus.Maj(ref F_Dessus);
                        //pidF_Devant.Maj(ref F_Devant);
                        //pidPltG_Plan.Maj(ref PltG_Plan);
                        //pidPltD_Plan.Maj(ref PltD_Plan);
                        //pidPltG.Maj(ref PltG);
                        //pidPltD.Maj(ref PltD);

                        pidF_Dessus.Maj();
                        pidF_Devant.Maj();
                        pidPltG_Plan.Maj();
                        pidPltD_Plan.Maj();
                        pidPltG.Maj();
                        pidPltD.Maj();

                        Run(F_Dessus, F_Devant, PltG_Plan, PltD_Plan, PltG, PltD);
                    }

                    MdlBase.ShowConfiguration2(_NomConfigCourante);
                }
                else
                {
                    Run(F_Dessus, F_Devant, PltG_Plan, PltD_Plan, PltG, PltD);
                }

                MdlBase.EditRebuild3();
            }

            private void Run(Face2 dessus, Face2 devant, Feature gPlan, Feature dPlan, Component2 pltG, Component2 pltD)
            {
                if ((dessus == null) || (devant == null))
                {
                    this.LogMethode(new String[] { "Une reference à un objet a été perdue" });
                    return;
                }
                try
                {
                    Edge E_Face = dessus.eListeDesArretesCommunes(devant)[0];

                    List<Edge> ListeArrete = dessus.eListeDesArretesContigues(E_Face);

                    // On assigne les cotes de façon arbitraire pour eviter une assignation suplémentaire
                    Edge E_Gauche = ListeArrete[0];
                    Edge E_Droit = ListeArrete[1];

                    // Création des segements
                    Segment S1 = new Segment(E_Gauche);
                    Segment Sf = new Segment(E_Face);

                    // Orientation des segements
                    S1.OrienterDe(Sf);
                    Sf.OrienterVers(S1);

                    Vecteur Normal = new Vecteur((Double[])dessus.Normal);

                    // Verification du sens de rotation et modification des cotes si nécessaire
                    if (Sf.Vecteur.RotationTrigo(S1.Vecteur, Normal))
                    {
                        E_Gauche = ListeArrete[1];
                        E_Droit = ListeArrete[0];
                    }
                    List<Face2> L = null;

                    L = E_Gauche.eListeDesFaces();
                    L.Remove(dessus);
                    Face2 FaceGauche = L[0];

                    L = E_Droit.eListeDesFaces();
                    L.Remove(dessus);
                    Face2 FaceDroite = L[0];

                    if (pltG.IsRef())
                        if (pltG.GetConstrainedStatus() == (int)swConstrainedStatus_e.swUnderConstrained)
                            Contraindre(gPlan, FaceGauche);

                    if (pltD.IsRef())
                        if (pltD.GetConstrainedStatus() == (int)swConstrainedStatus_e.swUnderConstrained)
                            Contraindre(dPlan, FaceDroite);
                }
                catch (Exception e)
                {
                    this.LogMethode(new Object[] { e });
                }
            }

            private void Contraindre(Feature plan, Face2 e)
            {
                AssemblyDoc Ass = MdlBase.eAssemblyDoc();

                e.eSelectEntite(MdlBase);
                plan.eSelect(true);

                int longstatus = 0;
                Ass.AddMate5((int)swMateType_e.swMateCOINCIDENT,
                             (int)swMateAlign_e.swMateAlignCLOSEST, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out longstatus);

                MdlBase.ClearSelection2(true);
            }
        }
    }
}


