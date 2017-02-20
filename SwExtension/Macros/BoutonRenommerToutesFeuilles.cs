using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Renommer toutes les feuilles"),
        ModuleNom("RenommerToutesFeuilles")]
    public class BoutonRenommerToutesFeuilles : BoutonBase
    {
        protected override void Command()
        {
            DrawingDoc Dessin = App.DrawingDoc;


            String PropDesignation = App.Sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swCustomPropertyUsedAsComponentDescription);

            int i = 1;
            foreach (Sheet feuille in Dessin.eListeDesFeuilles())
                feuille.SetName((i++).ToString());

            List<String> ListeNomFeuille = new List<string>();
            Dictionary<String, Sheet> DicFeuilles = new Dictionary<string, Sheet>();

            foreach (Sheet feuille in Dessin.eListeDesFeuilles())
            {
                var ListeVues = feuille.eListeDesVues();
                if (ListeVues.Count == 0) continue;

                View v = ListeVues[0];
                if (v.IsNull()) continue;

                ModelDoc2 mdlVue = v.ReferencedDocument;
                if (mdlVue.IsNull()) continue;

                String Designation = "";

                if (mdlVue.ePropExiste(PropDesignation))
                    Designation = mdlVue.eProp(PropDesignation);

                String NomFeuille =( mdlVue.eNomSansExt() + "-" + v.ReferencedConfiguration + " " + Designation).Trim();

                String NomTemp = NomFeuille;
                int Indice = 1;

                while (DicFeuilles.ContainsKey(NomTemp))
                {
                    if(Indice == 1)
                    {
                        Sheet f = DicFeuilles[NomTemp];
                        DicFeuilles.Remove(NomTemp);
                        NomTemp = NomFeuille + " (" + Indice++ + ")";
                        DicFeuilles.Add(NomTemp, f);
                    }
                    NomTemp = NomFeuille + " (" + Indice++ + ")";
                }

                NomFeuille = NomTemp;

                DicFeuilles.Add(NomTemp, feuille);
            }

            foreach (String NomFeuille in DicFeuilles.Keys)
            {
                Sheet feuille = DicFeuilles[NomFeuille];
                feuille.SetName(NomFeuille);
                WindowLog.Ecrire(NomFeuille);
            }

            App.ModelDoc2.ForceRebuild3(false);
        }
    }
}
