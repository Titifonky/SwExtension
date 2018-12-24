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

                AjouterInfos(MdlBase);

                if (MdlBase.TypeDoc() == eTypeDoc.Assemblage)
                    MdlBase.eRecParcourirComposants(c => { if (c.IsHidden(true)) return false; AjouterInfos(c.eModelDoc2()); return false; });

                if (ToutReconstruire)
                    MdlBase.ForceRebuild3(false);
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
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

        private void AjouterInfos(ModelDoc2 Mdl)
        {
            try
            {
                if (!_ListeComp.Contains(Mdl.GetPathName()) && (ComposantsExterne || Mdl.eEstDansLeDossier(MdlBase)))
                {
                    _ListeComp.Add(Mdl.GetPathName());

                    WindowLog.Ecrire(Mdl.eNomAvecExt());
                    CustomPropertyManager PM = Mdl.Extension.get_CustomPropertyManager("");
                    foreach (String k in _Dic.Keys)
                        PM.Add3(k, (int)swCustomInfoType_e.swCustomInfoText, _Dic[k], (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                }
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }
    }
}


