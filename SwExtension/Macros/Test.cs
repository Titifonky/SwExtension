using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using MS.Internal.IO.Packaging;
using System.Security;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Test"),
        ModuleNom("Test")]
    public class Test : BoutonBase
    {
        private Parametre Pid;

        public Test()
        {
            Pid = _Config.AjouterParam("Pid", "xxx");
        }

        protected override void Command()
        {
            try
            {
                var face = MdlBase.eSelect_RecupererObjet<Face2>();
                var corps = (Body2)face.GetBody();
                MemoryStream ms = new MemoryStream();
                ManagedIStream MgIs = new ManagedIStream(ms);
                var copie = (Body2)corps.Copy2(true);
                copie.Save(MgIs);
                var Tab = ms.ToArray();
                WindowLog.Ecrire(Tab.Length);
                File.WriteAllBytes(Path.Combine(MdlBase.eDossier(),"Corps.data") , Tab);
            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
            //Appliquer(MdlBase);
        }

        private void Appliquer(ModelDoc2 mdl)
        {
            String NomFichier = mdl.eNomAvecExt().Replace(".SLDLFP", ".SLDPRT");
            String Materiau = "\"SW-Material@@<NomCfg>@<NomFichier>\"".Replace("<NomFichier>", NomFichier);
            String Masse = "\"SW-Mass@@<NomCfg>@<NomFichier>\"".Replace("<NomFichier>", NomFichier);

            foreach (var NomCfg in mdl.eListeNomConfiguration())
            {
                String NomProfil = NomCfg;
                var tmp = NomCfg.Split(new char[] { 'x' }).ToList();
                if (tmp.Count > 1)
                {
                    tmp.RemoveAt(tmp.Count - 1);
                    NomProfil = String.Join("x", tmp);
                }
                mdl.ePropAdd("ProfilCourt", NomProfil, NomCfg);
                mdl.ePropAdd("Matériau", Materiau.Replace("<NomCfg>", NomCfg), NomCfg);
                mdl.ePropAdd("Masse", Masse.Replace("<NomCfg>", NomCfg), NomCfg);
            }
        }
    }
}
