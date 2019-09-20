using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using Outils;
using LogDebugging;
using SolidWorks.Interop.swconst;
using System.Linq;
using SwExtension;

namespace ModuleInsererPercageTole
{
    public class CmdInsererPercageTole : Cmd
    {
        private ModelDoc2 _MdlBase = null;
        private AssemblyDoc _AssBase = null;
        public ModelDoc2 MdlBase
        {
            get { return _MdlBase; }
            set { _MdlBase = value; _AssBase = value.eAssemblyDoc(); }
        }

        public Component2 CompBase = null;
        public Component2 CompPercage = null;
        public List<Double> ListeDiametre = new List<double>() { 0 };
        public Boolean PercageOuvert = false;

        
        private Dictionary<String, String> _ListePercage = new Dictionary<String, String>();

        private Dictionary<String, List<String>> _DicConfigWithComp = new Dictionary<string, List<String>>();

        protected override void Command()
        {
            try
            {
                App.Sw.CommandInProgress = true;
                MdlBase.eActiverManager(false);

                AjouterPercage(CompPercage);

                Run(CompBase);
                InsererDansUnDossier();

                // On met les percages dans un dossier, c'est plus propre

                _MdlBase.EditRebuild3();
                MdlBase.eActiverManager(true);
                App.Sw.CommandInProgress = false;
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        private void Run(Component2 cBase)
        {
            if (cBase == null)
            {
                this.LogMethode(new String[] { "Une reference à un objet a été perdue cBase :", cBase.IsRefToString() });
                return;
            }

            try
            {
                List<Body2> ListeCorps = cBase.eListeCorps();
                List<Trou> ListeTrou = new List<Trou>();

                // Recherche des faces cylindriques
                foreach (Body2 C in ListeCorps)
                {
                    Face2 faceBase = null;
                    foreach (Feature feature in (object[])C.GetFeatures())
                    {
                        if (feature.GetTypeName2() == FeatureType.swTnFlatPattern)
                        {
                            // On récupère la face fixe
                            var def = (FlatPatternFeatureData)feature.GetDefinition();
                            faceBase = (Face2)def.FixedFace2;

                            // On liste les faces du corps et on regarde 
                            // la face qui est la même que celle de la face fixe
                            foreach (var face in C.eListeDesFaces())
                            {
                                if (face.IsSame(faceBase))
                                {
                                    faceBase = face;
                                    break;
                                }
                            }
                            break;
                        }
                    }

                    foreach (Face2 F in C.eListeDesFaces())
                    {
                        Surface S = F.GetSurface();
                        if (S.IsCylinder() && (PercageOuvert || (F.GetLoopCount() > 1)))
                            {
                            // Si le diametre == 0 on recupère toutes les faces
                            if ((ListeDiametre.Count == 1) && (ListeDiametre[0] == 0))
                                ListeTrou.Add(new Trou(F, faceBase));
                            // Sinon, on verifie qu'elle corresponde bien au diametre demandé
                            else if (ListeDiametre.Contains(DiametreMm(S)))
                                ListeTrou.Add(new Trou(F, faceBase));
                        }
                    }
                }

                // S'il n'y a pas assez de composant de perçage, on en rajoute
                while (_ListePercage.Count < ListeTrou.Count)
                    AjouterPercage();

                _MdlBase.EditRebuild3();

                List<String> ListeNomComp = _ListePercage.Keys.ToList();

                // Contrainte des perçages
                for (int i = 0; i < ListeTrou.Count; i++)
                {
                    String NomComp = ListeNomComp[i];
                    Component2 Comp = _AssBase.GetComponentByName(NomComp);
                    // On active le composant au cas ou :)

                    _MdlBase.eSelectByIdComp(_ListePercage[NomComp]);
                    _AssBase.eModifierEtatComposant(swComponentSuppressionState_e.swComponentResolved);
                    _MdlBase.eEffacerSelection();

                    int longstatus = 0;
                    var Trou = ListeTrou[i];

                    _MdlBase.ClearSelection2(true);
                    Trou.Plan.eSelectEntite();
                    PlanDeFace(Comp).eSelect(true);
                    
                    _AssBase.AddMate5((int)swMateType_e.swMateCOINCIDENT,
                                    (int)swMateAlign_e.swMateAlignALIGNED,
                                    false,
                                    0, 0, 0, 0, 0, 0, 0, 0,
                                    false, false, 0, out longstatus);

                    _MdlBase.eEffacerSelection();
                    Trou.Cylindre.eSelectEntite();
                    FaceCylindrique(Comp).eSelectEntite(true);

                    _AssBase.AddMate5((int)swMateType_e.swMateCONCENTRIC,
                                    (int)swMateAlign_e.swMateAlignCLOSEST,
                                    false,
                                    0, 0, 0, 0, 0, 0, 0, 0,
                                    false, true, 0, out longstatus);

                    _MdlBase.eEffacerSelection();
                }

            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }

        private Double DiametreMm(Surface s)
        {
            Double[] ListeParam = (Double[])s.CylinderParams;
            return Math.Round(ListeParam[6] * 2.0 * 1000, 2);
        }

        private class Trou
        {
            public Face2 Cylindre;
            public Face2 Plan;
            public Trou(Face2 cylindre, Face2 plan) { Cylindre = cylindre; Plan = plan; }
        }

        /// <summary>
        /// Creer un percage ou ajouter
        /// </summary>
        /// <param name="comp"></param>
        private void AjouterPercage(Component2 comp = null)
        {
            Component2 Comp = comp;

            if (Comp.IsNull())
            {
                Comp = _AssBase.AddComponent5(CompPercage.GetPathName(),
                                            (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig,
                                            "",
                                            true,
                                            CompPercage.eNomConfiguration(),
                                            0, 0, 0);

            }

            _MdlBase.eSelectByIdComp(Comp.GetSelectByIDString());
            _AssBase.eModifierEtatComposant(swComponentSuppressionState_e.swComponentResolved);
            _MdlBase.eEffacerSelection();

            _ListePercage.Add(Comp.Name2, Comp.GetSelectByIDString());
        }

        private void InsererDansUnDossier()
        {

            if (_ListePercage.Count > 0)
            {
                _MdlBase.eEffacerSelection();

                foreach (String NomCompSelection in _ListePercage.Values)
                {
                    _MdlBase.eSelectByIdComp(NomCompSelection, -1, true);
                }

                Feature fDossier = _MdlBase.FeatureManager.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing);
                fDossier.Name = "Percage";

                _MdlBase.eEffacerSelection();
            }
        }

        // Renvoi la première face cylindrique du composant
        private Face2 FaceCylindrique(Component2 composant)
        {
            List<Body2> ListeCorps = composant.eListeCorps();

            List<Face2> ListeFace = new List<Face2>();

            foreach (Body2 C in ListeCorps)
            {
                foreach (Face2 F in C.eListeDesFaces())
                {
                    Surface S = F.GetSurface();
                    if (S.IsCylinder())
                        return F;
                }
            }

            return null;
        }

        // Renvoi le plan de contrainte "Plan de face"
        private Feature PlanDeFace(Component2 composant)
        {
            return composant.FeatureByName("Plan de face");
        }
    }
}


