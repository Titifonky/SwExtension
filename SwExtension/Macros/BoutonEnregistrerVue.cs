using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Enregistrer la vue"),
        ModuleNom("EnregistrerVue")]
    public class BoutonEnregistrerVue : BoutonBase
    {
        private Parametre PropVuesStandard;

        public BoutonEnregistrerVue()
        {
            LogToWindowLog = false;

            PropVuesStandard = _Config.AjouterParam("VuesStandard",
                                                    "*Normal à,*Face,*Arrière,*Gauche,*Droite,*Dessus,*Dessous,*Isométrique,*Trimétrique,*Dimétrique",
                                                    "Liste des vues standard à exclure (séparé par une virgule) :");
        }

        protected override void Command()
        {
            
            ModelDoc2 mdl = App.ModelDoc2;

            String nom = "";
            var hashVuesStandard = new HashSet<String>(PropVuesStandard.GetValeur<String>().Split(','));
            var listeVues = new List<String>();

            foreach (String n in mdl.GetModelViewNames())
            {
                if (!hashVuesStandard.Contains(n))
                    listeVues.Add(n);
            }

            if (listeVues.Count == 0)
                listeVues.Add("Aucunes");

            String VuesExistantes = "Vues existantes : " + String.Join(", ", listeVues);

            if (Interaction.InputBox("Enregistrer la vue", VuesExistantes, ref nom) == DialogResult.OK)
            {
                if (!String.IsNullOrWhiteSpace(nom) || !nom.StartsWith("*"))
                {
                    mdl.DeleteNamedView(nom);
                    mdl.NameView(nom);
                }
            }
        }
    }
}
