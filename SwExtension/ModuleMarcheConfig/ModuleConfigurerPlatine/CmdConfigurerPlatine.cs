using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace ModuleMarcheConfig
{
    namespace ModuleConfigurerPlatine
    {
        public class CmdConfigurerPlatine : Cmd
        {
            public ModelDoc2 MdlBase = null;
            public Face2 F_Dessus = null;
            public Component2 Marche = null;
            public Component2 PltG_Contrainte_Comp = null;
            public Component2 PltD_Contrainte_Comp = null;
            public Feature PltG_Contrainte_Plan = null;
            public Feature PltD_Contrainte_Plan = null;
            public Component2 PltG_Esquisse_Comp = null;
            public Component2 PltD_Esquisse_Comp = null;
            public Feature PltG_Esquisse_Fonction = null;
            public Feature PltD_Esquisse_Fonction = null;
            //public String NomEsquisse = "Config";
            public String Ag1 = "Ag1";
            public String Ag2 = "Ag2";
            public String Lg = "Lg";
            public Double LgMin = 0.08;
            public Boolean SurTouteLesConfigs = false;

            private String _NomConfigCourante = "";

            protected override void Command()
            {

                try
                {
                    _NomConfigCourante = MdlBase.eNomConfigActive();

                    if (SurTouteLesConfigs)
                    {
                        var pidF_Dessus = new SwObjectPID<Face2>(F_Dessus, MdlBase);
                        var pidMarche = new SwObjectPID<Component2>(Marche, MdlBase);
                        var pidPltG_Contrainte_Comp = new SwObjectPID<Component2>(PltG_Contrainte_Comp, MdlBase);
                        var pidPltD_Contrainte_Comp = new SwObjectPID<Component2>(PltD_Contrainte_Comp, MdlBase);
                        var pidPltG_Contrainte_Plan = new SwObjectPID<Feature>(PltG_Contrainte_Plan, MdlBase);
                        var pidPltD_Contrainte_Plan = new SwObjectPID<Feature>(PltD_Contrainte_Plan, MdlBase);

                        var pidPltG_Esquisse_Comp = new SwObjectPID<Component2>(PltG_Esquisse_Comp, MdlBase);
                        var pidPltD_Esquisse_Comp = new SwObjectPID<Component2>(PltD_Esquisse_Comp, MdlBase);
                        var pidPltG_Esquisse_Fonction = new SwObjectPID<Feature>(PltG_Esquisse_Fonction, MdlBase);
                        var pidPltD_Esquisse_Fonction = new SwObjectPID<Feature>(PltD_Esquisse_Fonction, MdlBase);

                        List<String> ListeNomsConfig = MdlBase.eListeNomConfiguration(eTypeConfig.DeBase);

                        String NomMarche = Marche.eNomSansExt();
                        String NomPltG = "";
                        if (PltG_Contrainte_Comp.IsRef())
                            NomPltG = PltG_Contrainte_Comp.eNomSansExt();

                        String NomPltD = "";
                        if (PltD_Contrainte_Comp.IsRef())
                            NomPltD = PltD_Contrainte_Comp.eNomSansExt();

                        Log.Message(NomPltG);
                        Log.Message(NomPltD);

                        foreach (String NomConfig in ListeNomsConfig)
                        {
                            MdlBase.ShowConfiguration2(NomConfig);
                            MdlBase.EditRebuild3();

                            pidF_Dessus.Maj(ref F_Dessus);
                            pidMarche.Maj(ref Marche);
                            pidPltG_Contrainte_Comp.Maj(ref PltG_Contrainte_Comp);
                            pidPltD_Contrainte_Comp.Maj(ref PltD_Contrainte_Comp);
                            pidPltG_Contrainte_Plan.Maj(ref PltG_Contrainte_Plan);
                            pidPltD_Contrainte_Plan.Maj(ref PltD_Contrainte_Plan);

                            pidPltG_Esquisse_Comp.Maj(ref PltG_Esquisse_Comp);
                            pidPltD_Esquisse_Comp.Maj(ref PltD_Esquisse_Comp);
                            pidPltG_Esquisse_Fonction.Maj(ref PltG_Esquisse_Fonction);
                            pidPltD_Esquisse_Fonction.Maj(ref PltD_Esquisse_Fonction);

                            if (PltG_Contrainte_Comp.IsRef())
                                Run(MdlBase, Marche, PltG_Contrainte_Comp, PltG_Contrainte_Plan, F_Dessus, NomConfig, PltG_Esquisse_Fonction);

                            pidF_Dessus.Maj(ref F_Dessus);
                            pidMarche.Maj(ref Marche);
                            pidPltG_Contrainte_Comp.Maj(ref PltG_Contrainte_Comp);
                            pidPltD_Contrainte_Comp.Maj(ref PltD_Contrainte_Comp);
                            pidPltG_Contrainte_Plan.Maj(ref PltG_Contrainte_Plan);
                            pidPltD_Contrainte_Plan.Maj(ref PltD_Contrainte_Plan);

                            pidPltG_Esquisse_Comp.Maj(ref PltG_Esquisse_Comp);
                            pidPltD_Esquisse_Comp.Maj(ref PltD_Esquisse_Comp);
                            pidPltG_Esquisse_Fonction.Maj(ref PltG_Esquisse_Fonction);
                            pidPltD_Esquisse_Fonction.Maj(ref PltD_Esquisse_Fonction);

                            if (PltD_Contrainte_Comp.IsRef())
                                Run(MdlBase, Marche, PltD_Contrainte_Comp, PltD_Contrainte_Plan, F_Dessus, NomConfig, PltD_Esquisse_Fonction);

                            MdlBase.EditRebuild3();
                        }
                    }
                    else
                    {
                        if (PltG_Contrainte_Comp.IsRef())
                            Run(MdlBase, Marche, PltG_Contrainte_Comp, PltG_Contrainte_Plan, F_Dessus, _NomConfigCourante, PltG_Esquisse_Fonction);

                        if (PltD_Contrainte_Comp.IsRef())
                            Run(MdlBase, Marche, PltD_Contrainte_Comp, PltD_Contrainte_Plan, F_Dessus, _NomConfigCourante, PltD_Esquisse_Fonction);
                    }

                    MdlBase.ShowConfiguration2(_NomConfigCourante);

                    MdlBase.EditRebuild3();

                }
                catch (Exception e)
                {
                    this.LogErreur(new Object[] { e });
                }
            }

            private void Run(ModelDoc2 modele, Component2 marche, Component2 platine, Feature planContrainte, Face2 f_Dessus, String nomConfig, Feature esquisse)
            {
                try
                {
                    Edge eBase = ArretePlatine(platine.eChercherContraintes(marche, false), planContrainte, f_Dessus);
                    Edge eFace = ArreteDevant(platine.eChercherContraintes(marche, false), f_Dessus);
                    Edge eArriere = ArreteArriere(f_Dessus, eBase, eFace);

                    gSegment F = new gSegment(eFace);
                    gSegment C = new gSegment(eBase);
                    gSegment A = new gSegment(eArriere);

                    C.OrienterDe(F);
                    A.OrienterDe(C);
                    F.OrienterDe(C);

                    Double gAg1 = F.Vecteur.Angle(C.Vecteur);
                    Double gAg2 = A.Vecteur.Angle(C.Vecteur.Inverse());

                    Double gLg = Math.Max(LgMin, C.Lg);

                    Configurer(platine, gAg1, gAg2, gLg, esquisse);
                }
                catch (Exception e)
                {
                    this.LogMethode(new Object[] { e });
                }
            }

            // Recherche de l'arrete sur laquelle est contrainte la platine.
            private Edge ArretePlatine(List<Mate2> listeContraintes, Feature planContrainte, Face2 f_Dessus)
            {
                Mate2 Contrainte = null;
                foreach (Mate2 Ct in listeContraintes)
                {
                    foreach (MateEntity2 Ent in Ct.eListeDesEntitesDeContrainte())
                    {
                        // Si l'entite est un plan
                        if (Ent.ReferenceType2 == (int)swSelectType_e.swSelDATUMPLANES)
                        {
                            RefPlane P = Ent.Reference;
                            Feature F = (Feature)P;
                            // On vérifie que le plan a le même nom que le plan contrainte
                            if (planContrainte.Name == F.Name)
                            {
                                Contrainte = Ct;
                                break;
                            }
                        }
                    }

                    if (Contrainte.IsRef()) break;
                }

                foreach (MateEntity2 Ent in Contrainte.eListeDesEntitesDeContrainte())
                {
                    // On récupère la face associée à la contrainte
                    if (Ent.ReferenceType2 == (int)swSelectType_e.swSelFACES)
                    {
                        Face2 F = Ent.Reference;

                        // Liste des arrêtes communes, normalement, il n'y en a qu'une
                        List<Edge> L = F.eListeDesArretesCommunes(f_Dessus);

                        // On renvoi la première
                        if (L.Count > 0)
                            return L[0];
                    }
                }

                return null;
            }

            // Recherche de l'arrete de devant sur laquelle est contrainte la platine.
            private Edge ArreteDevant(List<Mate2> listeContraintes, Face2 faceDessus)
            {
                Mate2 Ct = null;

                // On rechercher la contrainte avec un point
                foreach (Mate2 Contrainte in listeContraintes)
                {
                    foreach (MateEntity2 Ent in Contrainte.eListeDesEntitesDeContrainte())
                    {
                        if (Ent.ReferenceType2 == (int)swSelectType_e.swSelEXTSKETCHPOINTS)
                        {
                            Ct = Contrainte;
                            break;
                        }
                    }
                    if (Ct.IsRef()) break;
                }

                if (Ct.IsNull()) return null;

                Face2 F_Devant = null;

                // On recherche la face associée à cette contrainte
                foreach (MateEntity2 Ent in Ct.eListeDesEntitesDeContrainte())
                {
                    if (Ent.ReferenceType2 == (int)swSelectType_e.swSelFACES)
                    {
                        F_Devant = (Face2)Ent.Reference;
                        break;
                    }
                }

                if (F_Devant.IsNull()) return null;

                List<Edge> ListeArretes = F_Devant.eListeDesArretesCommunes(faceDessus);

                if (ListeArretes.Count > 0)
                    return ListeArretes[0];

                return null;
            }

            // Recherche de l'arrete de derrière
            private Edge ArreteArriere(Face2 faceDessus, Edge eBase, Edge eFace)
            {
                List<Edge> Liste = faceDessus.eListeDesArretesContigues(eBase);

                // On supprime l'arrete de face
                Liste.Remove(eFace);

                // Il n'en reste qu'une et normalement c'est la bonne
                if (Liste.Count > 0)
                    return Liste[0];

                return null;
            }

            private void Configurer(Component2 composant, Double ag1, Double ag2, Double lg, Feature esquisse)
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
                        else if (D.Name == Lg)
                            D.SetSystemValue3(lg, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

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


