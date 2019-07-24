using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Text;
using System.IO;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Test"),
        ModuleNom("Test")]
    public class Test : BoutonBase
    {
        public Test() { }

        protected override void Command()
        {
            try
            {
                Body2 CorpsBase = null;
                String MateriauCorpsBase = "";
                if (MdlBase.eSelect_RecupererTypeObjet() == e_swSelectType.swSelFACES)
                {
                    var face = MdlBase.eSelect_RecupererObjet<Face2>();
                    CorpsBase = face.GetBody();
                }
                else if (MdlBase.eSelect_RecupererTypeObjet() == e_swSelectType.swSelSOLIDBODIES)
                {
                    CorpsBase = MdlBase.eSelect_RecupererObjet<Body2>();
                }

                if (CorpsBase == null)
                {
                    System.Windows.Forms.MessageBox.Show("Erreur de corps selectionné");
                    return;
                }

                MateriauCorpsBase = CorpsBase.eGetMateriauCorpsOuPiece(MdlBase.ePartDoc(), MdlBase.eNomConfigActive());

                System.Windows.Forms.OpenFileDialog pDialogue = new System.Windows.Forms.OpenFileDialog
                {
                    Filter = "Fichier texte (*.data)|*.data|Tout les fichiers (*.*)|*.*",
                    Multiselect = false,
                    InitialDirectory = Path.GetDirectoryName(Path.Combine(MdlBase.GetPathName(), "Pieces", "Corps")),
                    RestoreDirectory = true
                };

                String Chemin = "";

                if (pDialogue.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    Chemin = pDialogue.FileName;
                else
                {
                    System.Windows.Forms.MessageBox.Show("Erreur de chemin");
                    return;
                }

                Body2 CorpsCharge = null;
                String MateriauCorpsCharge = "";
                String NoCorps = "";
                Byte[] Tab = File.ReadAllBytes(Chemin);
                using (MemoryStream ms = new MemoryStream(Tab))
                {
                    ManagedIStream MgIs = new ManagedIStream(ms);
                    Modeler mdlr = (Modeler)App.Sw.GetModeler();
                    CorpsCharge = (Body2)mdlr.Restore(MgIs);
                }

                if (CorpsCharge == null)
                {
                    System.Windows.Forms.MessageBox.Show("Erreur de corps chargé");
                    return;
                }

                NoCorps = Path.GetFileNameWithoutExtension(Chemin).Substring(1);

                System.Windows.Forms.MessageBox.Show("NoCorps : " + NoCorps);

                var cheminNomenclature = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Chemin)), "Nomenclature.txt");

                System.Windows.Forms.MessageBox.Show("Chemin Nomenclature : " + cheminNomenclature);

                using (var sr = new StreamReader(cheminNomenclature, Encoding.GetEncoding(1252)))
                {
                    // On lit la première ligne contenant l'entête des colonnes
                    String ligne = sr.ReadLine();

                    if (ligne.IsRef())
                    {
                        while ((ligne = sr.ReadLine()) != null)
                        {
                            if (!String.IsNullOrWhiteSpace(ligne))
                            {
                                var tab = ligne.Split(new char[] { '\t' });
                                if (NoCorps == tab[0])
                                {
                                    MateriauCorpsCharge = tab[4];
                                    break;
                                }
                            }
                        }
                    }
                }

                if (String.IsNullOrWhiteSpace(MateriauCorpsCharge))
                {
                    System.Windows.Forms.MessageBox.Show("Erreur de materiau corps chargé");
                    return;
                }

                var Result = "Est different";
                if (MateriauCorpsCharge == MateriauCorpsBase)
                    Result = "Est semblable";

                System.Windows.Forms.MessageBox.Show("Materiau : " + MateriauCorpsCharge + " " + MateriauCorpsBase + "  " + Result);

                Result = "Est different";
                if (CorpsBase.eComparerGeometrie(CorpsCharge) == Sw.Comparaison_e.Semblable)
                    Result = "Est semblable";

                System.Windows.Forms.MessageBox.Show("Corps : " + Result);
            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }
    }
}
