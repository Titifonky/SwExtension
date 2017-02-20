using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using Outils;
using LogDebugging;
using SolidWorks.Interop.swconst;
using System.Linq;
using System.Text.RegularExpressions;
using SwExtension;

namespace ModuleContraindreComposant
{
    public class CmdContraindreComposant : Cmd
    {
        public ModelDoc2 MdlBase = null;
        public Component2 CompBase = null;
        public List<Component2> ListeComposants = null;
        public Boolean FixerComposant = false;

        private AssemblyDoc _Ass = null;

        protected override void Command()
        {
            _Ass = MdlBase.eAssemblyDoc();

            foreach(Component2 cp in ListeComposants)
                Run(cp);

            MdlBase.EditRebuild3();

        }

        private void Run(Component2 Comp)
        {
            try
            {
                List<Feature> ListePlanRef;

                if (CompBase.IsNull())
                    ListePlanRef = MdlBase.eListeFonctions(f => { return f.GetTypeName2() == FeatureType.swTnRefPlane; });
                else
                    ListePlanRef = CompBase.eListeFonctions(f => { return f.GetTypeName2() == FeatureType.swTnRefPlane; });

                Log.Message(Comp.Name2);
                List<Feature> ListePlan = Comp.eListeFonctions(f => { return f.GetTypeName2() == FeatureType.swTnRefPlane; });

                Boolean EstFixe = Comp.IsFixed();
                Comp.eLiberer(_Ass);

                Mate2 FxContrainte = null;
                int Erreur = 0;

                for (int i = 0; i < Math.Min(3, Math.Min(ListePlanRef.Count, ListePlan.Count)); i++)
                {
                    Feature PlanRef = ListePlanRef[i];
                    Feature Plan = ListePlan[i];

                    PlanRef.eSelectionnerPMP(MdlBase, -1, false);
                    Plan.eSelectionnerPMP(MdlBase, -1, true);

                    FxContrainte = _Ass.AddMate5((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, true, 0, out Erreur);
                }

                MdlBase.eEffacerSelection();

                if (EstFixe || FixerComposant)
                    Comp.eFixer(_Ass);
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }
    }
}


