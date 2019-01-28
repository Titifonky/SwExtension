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
        ModuleTitre("Afficher/Masquer l'esquisse de reperage"),
        ModuleNom("BoutonAfficherEsquisseAssemblage")]
    public class BoutonAfficherEsquisseAssemblage : BoutonBase
    {
        Parametre AfficherMasquer;
        Parametre NomEsquisse;

        public BoutonAfficherEsquisseAssemblage()
        {
            LogToWindowLog = false;
            AfficherMasquer = _Config.AjouterParam("AfficherMasquer", true, "Afficher/Masquer les esquisses");
            NomEsquisse = _Config.AjouterParam("NomEsquisse", "E - Reperage Dessin", "Nom de l'esquisse à modifier");
        }

        protected override void Command()
        {
            try
            {
                var ListeVue = MdlBase.eSelect_RecupererListeObjets<View>();

                if (ListeVue.IsNull() || (ListeVue.Count == 0))
                    ListeVue = MdlBase.eDrawingDoc().eFeuilleActive().eListeDesVues();

                if (ListeVue.IsNull() || (ListeVue.Count == 0)) return;

                MdlBase.eEffacerSelection();

                foreach (var vue in ListeVue)
                    ParcourirDrawingComponent(vue.RootDrawingComponent, vue, NomEsquisse.GetValeur<String>(), AfficherMasquer.GetValeur<Boolean>());

                MdlBase.eEffacerSelection();

                AfficherMasquer.SetValeur<Boolean>(!AfficherMasquer.GetValeur<Boolean>());
                _Config.Sauver();

            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        private void ParcourirDrawingComponent(DrawingComponent dcp, View vue, String nomEsquisse, Boolean afficher)
        {
            Component2 cp = dcp.Component;
            cp.eParcourirFonctions(
                f =>
                {
                    if (f.Name == nomEsquisse)
                    {
                        f.eSelectionnerById2Dessin(MdlBase, vue);
                        if (afficher)
                            MdlBase.UnblankSketch();
                        else
                            MdlBase.BlankSketch();

                        return true;
                    }
                    
                    return false;
                }
                , false);

            if (dcp.GetChildrenCount() > 0)
            {
                Object[] l = (object[])dcp.GetChildren();
                foreach (DrawingComponent sdcp in l)
                {
                    if (sdcp.Visible)
                        ParcourirDrawingComponent(sdcp, vue, nomEsquisse, afficher);
                }
            }
        }
    }
}
