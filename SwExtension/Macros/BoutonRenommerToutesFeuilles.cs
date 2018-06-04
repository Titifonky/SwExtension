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
            try
            {
                DrawingDoc Dessin = App.DrawingDoc;


                String PropDesignation = App.Sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swCustomPropertyUsedAsComponentDescription);

                int i = 1;
                foreach (Sheet feuille in Dessin.eListeDesFeuilles())
                    feuille.SetName((i++).ToString());

                List<String> ListeNomFeuille = new List<string>();
                Dictionary<String, List<Sheet>> DicFeuilles = new Dictionary<string, List<Sheet>>();

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

                    String NomFeuille = (mdlVue.eNomSansExt() + "-" + v.ReferencedConfiguration + " " + Designation).Trim();

                    String NomTemp = NomFeuille;

                    if (DicFeuilles.ContainsKey(NomTemp))
                    {
                        var l = DicFeuilles[NomTemp];
                        l.Add(feuille);
                    }
                    else
                    {
                        var l = new List<Sheet>() { feuille };
                        DicFeuilles.Add(NomTemp, l);
                    }
                }

                foreach (String NomFeuille in DicFeuilles.Keys)
                {
                    var l = DicFeuilles[NomFeuille];
                    if (l.Count > 1)
                    {
                        int j = 1;
                        foreach (var feuille in DicFeuilles[NomFeuille])
                        {
                            var n = NomFeuille + "(" + j++ + ")";
                            feuille.SetName(n);
                            WindowLog.Ecrire(n);
                        }
                    }
                    else
                    {
                        l[0].SetName(NomFeuille);
                        WindowLog.Ecrire(NomFeuille);
                    }
                }

                App.ModelDoc2.ForceRebuild3(false);
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }
    }
}
