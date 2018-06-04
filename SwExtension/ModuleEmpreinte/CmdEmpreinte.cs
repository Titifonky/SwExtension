using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using Outils;
using LogDebugging;
using SolidWorks.Interop.swconst;
using System.Linq;
using System.Text.RegularExpressions;
using SwExtension;

namespace ModuleEmpreinte
{
    public class CmdEmpreinte : Cmd
    {
        public ModelDoc2 MdlBase = null;
        public List<Component2> ListeCompBase = new List<Component2>();
        public List<Component2> ListeCompEmpreinte = new List<Component2>();

        private AssemblyDoc Ass = null;

        protected override void Command()
        {
            try
            {
                MdlBase.eEffacerSelection();
                Ass = MdlBase.eAssemblyDoc();

                Dictionary<Component2, Dictionary<String, List<Component2>>> Dic = new Dictionary<Component2, Dictionary<String, List<Component2>>>();

                foreach (Component2 cpB in ListeCompBase)
                {
                    foreach (Body2 cB in cpB.eListeCorps())
                    {
                        foreach (Component2 cpE in ListeCompEmpreinte)
                        {
                            List<Body2> l = cpE.eListeCorps();

                            if (l.IsRef() && (l.Count > 0))
                            {
                                Boolean t = l.Any(cE =>
                                                    {
                                                        if (cpB.eNbIntersection(cB, cpE, cE) > 0)
                                                        {
                                                            if (!Dic.ContainsKey(cpB))
                                                                Dic.Add(cpB, new Dictionary<string, List<Component2>>());

                                                            String nomEmpreinte = cpE.ValProp() + "_" + Empreinte.NOM_FONCTION;

                                                            if (!Dic[cpB].ContainsKey(nomEmpreinte))
                                                                Dic[cpB].Add(nomEmpreinte, new List<Component2>());

                                                            Dic[cpB][nomEmpreinte].Add(cpE);
                                                            return true;
                                                        }

                                                        return false;
                                                    }
                                                    );
                                if (t) continue;
                            }
                        }
                    }
                }


                foreach (Component2 cpB in Dic.Keys)
                {
                    MdlBase.eEffacerSelection();
                    Ass.eEditerLeComposant(cpB);

                    foreach (String nomEmpreinte in Dic[cpB].Keys)
                    {
                        foreach (Component2 cpE in Dic[cpB][nomEmpreinte])
                            cpE.eSelect(true);

                        ModelDoc2 mdl = cpB.eModelDoc2();
                        Feature f = mdl.Extension.GetLastFeatureAdded();

                        Ass.InsertCavity4(0, 0, 0, true, (int)swCavityScaleType_e.swAboutOrigin, -1);

                        Feature fEmpreinte = mdl.Extension.GetLastFeatureAdded();

                        if (fEmpreinte.IsRef() && (f.Name != fEmpreinte.Name))
                        {
                            fEmpreinte.eRenommerFonction(nomEmpreinte + "_");
                            WindowLog.Ecrire(fEmpreinte.Name);
                        }
                    }

                    Ass.eEditerAssemblage();

                }
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }

        }
    }
}


