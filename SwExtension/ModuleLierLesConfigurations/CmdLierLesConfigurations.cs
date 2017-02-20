using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using Outils;
using LogDebugging;
using SolidWorks.Interop.swconst;
using System.Linq;
using System.Text.RegularExpressions;
using SwExtension;

namespace ModuleLierLesConfigurations
{
    public class CmdLierLesConfigurations : Cmd
    {
        public ModelDoc2 MdlBase = null;
        public List<Component2> ListeComposants = null;
        public Boolean CreerConfigsManquantes = false;
        public List<String> ListeConfig = null;
        public Boolean SupprimerNvlFonction = true;
        public Boolean SupprimerNvComposant = true;

        private String NomConfigCourante = "";

        private int Errors = 0;
        private int Warnings = 0;

        protected override void Command()
        {
            try
            {
                NomConfigCourante = MdlBase.eNomConfigActive();

                // Si aucun composant n'a été selectionné, on sort
                if ((MdlBase.TypeDoc() != eTypeDoc.Assemblage) || (ListeComposants.IsNullOrEmpty())) return;

                List<List<String>> Liste = new List<List<string>>();

                foreach (Component2 cp in ListeComposants)
                {
                    var ListeCp = cp.eListeComposantParent();
                    ListeCp.Reverse();
                    ListeCp.Add(cp);

                    for (int i = 0; i < ListeCp.Count; i++)
                    {
                        if(i < Liste.Count)
                        {
                            var l = Liste[i];
                            l.AddIfNotExist(ListeCp[i].GetSelectByIDString());
                        }
                        else
                        {
                            var l = new List<String>() { ListeCp[i].GetSelectByIDString() };
                            Liste.Add(l);
                        }
                    }
                }

                CreerConfigs();

                MdlBase.EditRebuild3();
                MdlBase.Save3((int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced + (int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref Errors, ref Warnings);

                foreach (String nomConfig in ListeConfig)
                {
                    WindowLog.Ecrire("Config : \"" + nomConfig + "\"");
                    MdlBase.ShowConfiguration2(nomConfig);
                    MdlBase.EditRebuild3();

                    int i = 0;
                    foreach (var lNiveau in Liste)
                    {
                        i++;
                        foreach (var IdString in lNiveau)
                        {
                            MdlBase.eSelectByIdComp(IdString, 1, false);
                            Component2 cp = MdlBase.eSelect_RecupererComposant(1, 1);
                            cp.ReferencedConfiguration = nomConfig;
                            WindowLog.Ecrire( " ".eRepeter(i * 2) +  "- " + cp.Name2.Split('/').Last() + " \"" + cp.eNomConfiguration() + "\"");
                        }
                    }
                }

                MdlBase.ShowConfiguration2(NomConfigCourante);
                MdlBase.EditRebuild3();
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private Boolean CreerConfigs()
        {
            Dictionary<String, ModelDoc2> ListeMdl = new Dictionary<String, ModelDoc2>() { { MdlBase.GetPathName(), MdlBase } };
            Dictionary<String, String> ListeMdlDisplayState = new Dictionary<string, string>() { { MdlBase.GetPathName(), MdlBase.eComposantRacine().ReferencedDisplayState } };

            foreach (Component2 cp in ListeComposants)
            {
                ModelDoc2 mdl = cp.GetModelDoc2();
                ListeMdl.AddIfNotExist(mdl.GetPathName(), mdl);
                ListeMdlDisplayState.AddIfNotExist(mdl.GetPathName(), cp.ReferencedDisplayState);

                foreach (Component2 cpParent in cp.eListeComposantParent())
                {
                    ModelDoc2 mdlParent = cpParent.GetModelDoc2();
                    ListeMdl.AddIfNotExist(mdlParent.GetPathName(), mdlParent);
                    ListeMdlDisplayState.AddIfNotExist(mdlParent.GetPathName(), cpParent.ReferencedDisplayState);
                }
            }

            int Options = (int)swConfigurationOptions2_e.swConfigOption_InheritProperties;

            if (SupprimerNvlFonction)
                Options += (int)swConfigurationOptions2_e.swConfigOption_SuppressByDefault;

            if (SupprimerNvComposant)
                Options += (int)swConfigurationOptions2_e.swConfigOption_HideByDefault;

            if ((ListeConfig.IsRef()) && (ListeConfig.Count > 0) && !String.IsNullOrWhiteSpace(ListeConfig[0]))
            {
                foreach (ModelDoc2 mdl in ListeMdl.Values)
                {
                    WindowLog.Ecrire(mdl.eNomAvecExt());
                    WindowLog.Ecrire(" - " + String.Join(" ", ListeConfig));
                    foreach (String nomConfig in ListeConfig)
                    {
                        String NomCurrentDisplayState = ListeMdlDisplayState[mdl.GetPathName()];

                        mdl.eAddConfiguration(nomConfig, nomConfig, "", Options);
                    }

                    WindowLog.SautDeLigne();
                }

                return true;
            }

            return false;
        }
    }
}


