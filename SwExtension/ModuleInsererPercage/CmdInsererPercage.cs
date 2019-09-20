using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using Outils;
using LogDebugging;
using SolidWorks.Interop.swconst;
using System.Linq;
using SwExtension;

namespace ModuleInsererPercage
{
    public class CmdInsererPercage : Cmd
    {
        private ModelDoc2 _MdlBase;
        private AssemblyDoc _AssBase = null;
        public ModelDoc2 MdlBase
        {
            get { return _MdlBase; }
            set { _MdlBase = value; _AssBase = value.eAssemblyDoc(); }
        }

        public Component2 CompBase = null;
        public Component2 CompPercage = null;
        public Face2 Face = null;
        public Feature Plan = null;
        public List<Double> ListeDiametre = new List<double>() { 0 };
        public Boolean PercageOuvert = false;
        public Boolean SurTouteLesConfigs = false;

        
        private String _NomConfigCourante = "";
        private String _NomConfigActive = "";

        private Dictionary<String, String> _ListePercage = new Dictionary<String, String>();

        private Dictionary<String, List<String>> _DicConfigWithComp = new Dictionary<string, List<String>>();

        protected override void Command()
        {
            try
            {
                AjouterPercage(CompPercage);

                if (SurTouteLesConfigs)
                {
                    _NomConfigCourante = _MdlBase.ConfigurationManager.ActiveConfiguration.Name;

                    var pidBase = new SwObjectPID<Component2>(CompBase, _MdlBase);
                    var pidFace = new SwObjectPID<Face2>(Face, _MdlBase);
                    var pidPlan = new SwObjectPID<Feature>(Plan, _MdlBase);

                    List<String> ListeNomsConfig = _MdlBase.eListeNomConfiguration(eTypeConfig.DeBase);

                    foreach (String NomConfig in ListeNomsConfig)
                    {
                        _MdlBase.ShowConfiguration2(NomConfig);
                        _MdlBase.EditRebuild3();

                        Configuration Conf = _MdlBase.GetConfigurationByName(NomConfig);
                        Conf.SuppressNewFeatures = true;
                        //Conf.eSetSupprimerNouvellesFonctions(true, MdlBase);
                    }

                    foreach (String NomConfig in ListeNomsConfig)
                    {
                        _NomConfigActive = NomConfig;
                        _MdlBase.ShowConfiguration2(NomConfig);
                        _MdlBase.EditRebuild3();

                        pidBase.Maj(ref CompBase);
                        pidFace.Maj(ref Face);
                        pidPlan.Maj(ref Plan);

                        Run(CompBase);
                    }

                    InsererDansUnDossier();

                    foreach (String NomConfig in ListeNomsConfig)
                    {
                        _NomConfigActive = NomConfig;
                        _MdlBase.ShowConfiguration2(NomConfig);
                        _MdlBase.EditRebuild3();

                        List<String> ListeComp = _DicConfigWithComp[NomConfig];
                        foreach (String NomComp in _ListePercage.Keys)
                        {
                            if (!ListeComp.Contains(NomComp))
                            {
                                _MdlBase.eSelectByIdComp(_ListePercage[NomComp]);
                                _AssBase.eModifierEtatComposant(swComponentSuppressionState_e.swComponentSuppressed);
                                _MdlBase.eEffacerSelection();
                            }
                        }
                    }

                    _MdlBase.ShowConfiguration2(_NomConfigCourante);
                }
                else
                {
                    Run(CompBase);
                    InsererDansUnDossier();
                }

                // On met les percages dans un dossier, c'est plus propre

                _MdlBase.EditRebuild3();
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
                this.LogMethode(new String[]{_NomConfigActive});
                this.LogMethode(new String[] { "Une reference à un objet a été perdue cBase :" , cBase.IsRefToString() });
                return;
            }

            try
            {
                List<Body2> ListeCorps = cBase.eListeCorps();
                List<Face2> ListeFace = new List<Face2>();

                // Recherche des faces cylindriques
                foreach (Body2 C in ListeCorps)
                {
                    foreach (Face2 F in C.eListeDesFaces())
                    {
                        Surface S = F.GetSurface();
                        if (S.IsCylinder())
                        {
                            if (PercageOuvert || (F.GetLoopCount() > 1))
                            {
                                // Si le diametre == 0 on recupère toutes les faces
                                if ((ListeDiametre.Count == 1) && (ListeDiametre[0] == 0))
                                {
                                    if((Face == null) || (F.eListeDesFacesContigues().Contains(Face)))
                                        ListeFace.Add(F);
                                }
                                // Sinon, on verifie qu'elle corresponde bien au diametre demandé
                                else
                                {
                                    Double[] ListeParam = (Double[])S.CylinderParams;
                                    Double Diam = Math.Round(ListeParam[6] * 2.0 * 1000, 2);

                                    if (ListeDiametre.Contains(Diam))
                                        if ((Face == null) || (F.eListeDesFacesContigues().Contains(Face)))
                                            ListeFace.Add(F);

                                }
                            }
                        }
                    }
                }

                // S'il n'y a pas assez de composant de perçage, on en rajoute
                while (_ListePercage.Count < ListeFace.Count)
                    AjouterPercage();

                _MdlBase.EditRebuild3();

                // Mise à jour des références des perçages
                //ReinitialiserRefListe(ref _ListePercage);

                List<String> ListeComp = new List<String>();
                if (!_DicConfigWithComp.ContainsKey(_NomConfigActive))
                    _DicConfigWithComp.Add(_NomConfigActive, ListeComp);
                else
                    ListeComp = _DicConfigWithComp[_NomConfigActive];


                List<String> ListeNomComp = _ListePercage.Keys.ToList();

                // Contrainte des perçages
                for (int i = 0; i < ListeFace.Count; i++)
                {
                    String NomComp = ListeNomComp[i];
                    Component2 Comp = _AssBase.GetComponentByName(NomComp);
                    // On active le composant au cas ou :)

                    _MdlBase.eSelectByIdComp(_ListePercage[NomComp]);
                    _AssBase.eModifierEtatComposant(swComponentSuppressionState_e.swComponentResolved);
                    _MdlBase.eEffacerSelection();

                    Face2 Cylindre = ListeFace[i];

                    ListeComp.Add(NomComp);

                    //DesactiverContrainte(Comp);

                    _MdlBase.ClearSelection2(true);

                    if (Plan.IsRef())
                        Plan.eSelect();
                    else if (Face.IsRef())
                        Face.eSelectEntite();
                    else
                    {
                        Face2 FaceContrainte = FacePlane(Cylindre);

                        // Si on trouve une face pour la contrainte, on sélectionne
                        // Sinon, on passe au trou suivant
                        if (FaceContrainte.IsRef())
                            FaceContrainte.eSelectEntite();
                        else
                            continue;

                    }

                    Feature fPlan = PlanContrainte(Comp);

                    if (fPlan == null)
                    {
                        this.LogMethode(new String[]{"Pas de \"Plan de face\""});
                        continue;
                    }

                    fPlan.eSelect(true);

                    int longstatus = 0;
                    _AssBase.AddMate5((int)swMateType_e.swMateCOINCIDENT,
                                    (int)swMateAlign_e.swMateAlignCLOSEST,
                                    false,
                                    0, 0, 0, 0, 0, 0, 0, 0,
                                    false, false, 0, out longstatus);

                    _MdlBase.eEffacerSelection();

                    Cylindre.eSelectEntite();

                    Face2 fFace = FaceCylindrique(Comp);

                    fFace.eSelectEntite(true);

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
                this.LogMethode(new Object[]{e});
            }
        }

        /// <summary>
        /// Creer un percage ou ajouter
        /// </summary>
        /// <param name="comp"></param>
        private void AjouterPercage(Component2 comp = null)
        {
            Component2 Comp = comp;

            if (comp.IsNull())
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

        private void DesactiverContrainte(Component2 cp)
        {
            List<Mate2> Liste = cp.eListeContraintes(true);

            foreach (Mate2 M in Liste)
            {
                Feature F = M as Feature;
                F.eSelectionnerById2(_MdlBase);
                _MdlBase.EditSuppress2();
            }

            _MdlBase.ClearSelection2(true);
        }

        private void InsererDansUnDossier()
        {

            if (_ListePercage.Count > 0)
            {
                _MdlBase.eEffacerSelection();

                foreach (String NomCompSelection in _ListePercage.Values)
                {
                    _MdlBase.eSelectByIdComp(NomCompSelection,-1, true);
                }

                Feature fDossier = _MdlBase.FeatureManager.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing);
                fDossier.Name = "Percage";

                _MdlBase.eEffacerSelection();
            }
        }

        // Recherche d'une face plane attenante à la face
        private Face2 FacePlane(Face2 face)
        {
            List<Face2> ListeFaceTrou = face.eListeDesFacesContigues();
            foreach (Face2 F in ListeFaceTrou)
            {
                Surface S = F.GetSurface();

                if (S.IsPlane())
                    return F;
            }

            return null;
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
        private Feature PlanContrainte(Component2 composant)
        {
            return composant.FeatureByName("Plan de face");
            //return composant.eChercherFonction("Plan de face", false);
        }
    }
}


