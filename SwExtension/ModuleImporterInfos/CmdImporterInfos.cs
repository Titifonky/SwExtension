using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace ModuleImporterInfos
{
    public class CmdImporterInfos : Cmd
    {
        public ModelDoc2 MdlBase = null;
        public List<String> ListeValeurs = new List<String>();
        public Boolean ComposantsExterne = false;
        public Boolean ToutReconstruire = false;

        private Dictionary<String, String> _Dic = new Dictionary<String, String>();
        private HashSet<String> _ListeComp = new HashSet<String>();

        protected override void Command()
        {
            try
            {
                LireInfos();

                AjouterInfos(MdlBase.eComposantRacine());

                if (MdlBase.TypeDoc() == eTypeDoc.Assemblage)
                    MdlBase.eRecParcourirComposants(AjouterInfos);

                if (ToutReconstruire)
                    MdlBase.ForceRebuild3(false);
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private void LireInfos()
        {
            foreach (String v in ListeValeurs)
            {
                WindowLog.Ecrire(v);
                String[] l = v.Split(new Char[] { ':' }, 2, StringSplitOptions.None);
                _Dic.Add(l[0].Trim(), l[1].Trim());
            }

            WindowLog.SautDeLigne();
        }

        private Boolean AjouterInfos(Component2 Cp)
        {
            try
            {
                if (!_ListeComp.Contains(Cp.GetPathName()) && (ComposantsExterne || Cp.eEstDansLeDossier(MdlBase)))
                {
                    _ListeComp.Add(Cp.GetPathName());

                    WindowLog.Ecrire(Cp.eNomAvecExt());
                    CustomPropertyManager PM = Cp.eModelDoc2().Extension.get_CustomPropertyManager("");
                    foreach (String k in _Dic.Keys)
                        PM.Add3(k, (int)swCustomInfoType_e.swCustomInfoText, _Dic[k], (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                }
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }

            return false;
        }
    }
}


