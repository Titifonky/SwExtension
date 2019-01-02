using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace ModuleMarcheConfig
{
    namespace ModuleConfigurerContreMarche
    {
        public class CmdConfigurerContreMarche : Cmd
        {
            public ModelDoc2 MdlBase = null;
            public Face2 F_Dessus = null;
            public Face2 F_Devant = null;
            public Component2 ContreMarche_Esquisse_Comp = null;
            public Feature ContreMarche_Esquisse_Fonction = null;
            public Boolean SurTouteLesConfigs = false;
            public String NomEsquisse = "Config";
            public String Ag1 = "Ag1";
            public String Ag2 = "Ag2";
            public String Lg1 = "Lg1";
            public String Lg2 = "Lg2";
            public String Lc1 = "Lc1";
            public String Lc2 = "Lc2";

            private String _NomConfigCourante = "";

            protected override void Command()
            {
                try
                {
                    _NomConfigCourante = MdlBase.eNomConfigActive();

                    if (SurTouteLesConfigs)
                    {
                        var pidF_Dessus = new SwObjectPID<Face2>(F_Dessus, MdlBase);
                        var pidF_Devant = new SwObjectPID<Face2>(F_Devant, MdlBase);
                        var pidContreMarche_Esquisse_Comp = new SwObjectPID<Component2>(ContreMarche_Esquisse_Comp, MdlBase);
                        var pidContreMarche_Esquisse_Fonction = new SwObjectPID<Feature>(ContreMarche_Esquisse_Fonction, MdlBase);

                        List<String> ListeNomsConfig = MdlBase.eListeNomConfiguration(eTypeConfig.DeBase);

                        foreach (String NomConfig in ListeNomsConfig)
                        {
                            MdlBase.ShowConfiguration2(NomConfig);
                            MdlBase.EditRebuild3();

                            pidF_Dessus.Maj(ref F_Dessus);
                            pidF_Devant.Maj(ref F_Devant);
                            pidContreMarche_Esquisse_Comp.Maj(ref ContreMarche_Esquisse_Comp);
                            pidContreMarche_Esquisse_Fonction.Maj(ref ContreMarche_Esquisse_Fonction);

                            Run(F_Dessus, F_Devant, ContreMarche_Esquisse_Comp, ContreMarche_Esquisse_Fonction);
                        }

                        MdlBase.ShowConfiguration2(_NomConfigCourante);
                    }
                    else
                    {
                        Run(F_Dessus, F_Devant, ContreMarche_Esquisse_Comp, ContreMarche_Esquisse_Fonction);
                    }

                    MdlBase.EditRebuild3();
                }
                catch (Exception e)
                {
                    this.LogErreur(new Object[] { e });
                }
            }

            private void Run(Face2 dessus, Face2 devant, Component2 contreMarche, Feature esquisse)
            {
                if ((dessus == null) || (devant == null))
                {
                    this.LogMethode(new String[] { "Une reference à un objet a été perdue dessus | devant :", dessus.IsRefToString(), "|", devant.IsRefToString() });
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
                    gSegment S1 = new gSegment(E_Gauche);
                    gSegment Sf = new gSegment(E_Face);

                    // Orientation des segements
                    S1.OrienterDe(Sf);
                    Sf.OrienterVers(S1);

                    gVecteur Normal = new gVecteur((Double[])dessus.Normal);

                    // Verification du sens de rotation et modification des cotes si nécessaire
                    if (Sf.Vecteur.RotationTrigo(S1.Vecteur, Normal))
                    {
                        E_Gauche = ListeArrete[1];
                        E_Droit = ListeArrete[0];
                    }

                    gSegment F = new gSegment(E_Face);
                    gSegment G = new gSegment(E_Gauche);
                    gSegment D = new gSegment(E_Droit);

                    G.OrienterDe(F);
                    D.OrienterDe(F);
                    F.OrienterDe(G);

                    Double gAg1 = G.Vecteur.Angle(F.Vecteur);
                    Double gAg2 = D.Vecteur.Angle(F.Vecteur.Inverse());

                    Double gLg1 = new gPoint(0, 0, 0).Distance(F.Start);
                    Double gLg2 = new gPoint(0, 0, 0).Distance(F.End);
                    Double gLc1 = G.Lg;
                    Double gLc2 = D.Lg;

                    Configurer(contreMarche, gAg1, gAg2, gLg1, gLg2, gLc1, gLc2, esquisse);
                }
                catch (Exception e)
                {
                    this.LogErreur(new Object[] { e });
                }
            }

            private void Configurer(Component2 composant, Double ag1, Double ag2, Double lg1, Double lg2, Double lc1, Double lc2, Feature esquisse)
            {
                if (esquisse.IsRef())
                {
                    DisplayDimension DispDim = (DisplayDimension)esquisse.GetFirstDisplayDimension();
                    while (DispDim.IsRef())
                    {
                        Dimension D = (Dimension)DispDim.GetDimension();

                        if (D.Name == Ag1)
                            D.SetSystemValue3(ag1, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                        else if (D.Name == Ag2)
                            D.SetSystemValue3(ag2, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                        else if (D.Name == Lg1)
                            D.SetSystemValue3(lg1, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                        else if (D.Name == Lg2)
                            D.SetSystemValue3(lg2, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                        else if (D.Name == Lc1)
                            D.SetSystemValue3(lc1, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                        else if (D.Name == Lc2)
                            D.SetSystemValue3(lc2, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                        DispDim = (DisplayDimension)esquisse.GetNextDisplayDimension(DispDim);
                    }
                }
                else
                {
                    this.LogMethode(new String[] { "Ne trouve pas la fonction d'esquisse" });
                }
            }
        }
    }
}


