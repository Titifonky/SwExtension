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
                {
                    var dcp = vue.RootDrawingComponent;
                    if(dcp.IsRef())
                        ParcourirDrawingComponent(MdlBase, dcp, vue, NomEsquisse.GetValeur<String>(), AfficherMasquer.GetValeur<Boolean>());
                }

                MdlBase.eEffacerSelection();

                AfficherMasquer.SetValeur<Boolean>(!AfficherMasquer.GetValeur<Boolean>());
                _Config.Sauver();

            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        private void ParcourirDrawingComponent(ModelDoc2 mdlBase, DrawingComponent dcp, View vue, String nomEsquisse, Boolean afficher)
        {
            // Si le composant est Racine, on passe par le modele pour parcourir les fonctions
            // sinon ça ne marche pas
            if(dcp.Component.IsRoot())
                dcp.View.ReferencedDocument.eParcourirFonctions(f => AppliquerOptions(mdlBase, f, vue, nomEsquisse, afficher), false);
            else
                dcp.Component.eParcourirFonctions(f => AppliquerOptions(mdlBase, f, vue, nomEsquisse, afficher), false);

            try
            {
                if (dcp.GetChildrenCount() > 0)
                {
                    Object[] l = (object[])dcp.GetChildren();
                    foreach (DrawingComponent sdcp in l)
                    {
                        if (sdcp.Visible)
                            ParcourirDrawingComponent(mdlBase, sdcp, vue, nomEsquisse, afficher);
                    }
                }
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        private Boolean AppliquerOptions(ModelDoc2 mdlBase, Feature f, View vue, String nomEsquisse, Boolean afficher)
        {
            if (f.Name == nomEsquisse)
            {
                f.eSelectionnerById2Dessin(mdlBase, vue);
                if (afficher)
                    mdlBase.UnblankSketch();
                else
                    mdlBase.BlankSketch();

                return true;
            }

            return false;
        }
    }
}
