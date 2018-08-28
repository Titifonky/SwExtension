using LogDebugging;
using Outils;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;

namespace ModuleImporterInfos
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Importer des infos"),
        ModuleNom("ImporterInfos"),
        ModuleDescription("Importe les propriétés du fichier sélectionné dans tous les composants du modèle actif." + 
                          " Le fichier doit contenir une ligne par propriété, chaque propriété est definie comme ceci Nom : Valeur")
        ]
    public class PageImporterInfos : BoutonPMPManager
    {
        private Parametre _pFichierInfos;
        private Parametre _pComposantsExterne;
        private Parametre _pToutReconstruire;

        public PageImporterInfos()
        {
            try
            {
                _pFichierInfos = _Config.AjouterParam("FichierInfo", "Infos.txt", "Fichier à importer");
                _pComposantsExterne = _Config.AjouterParam("ComposantExterne", false, "Importer dans les composants externe au dossier du modèle");
                _pToutReconstruire = _Config.AjouterParam("ToutReconstruire", false, "Reconstruire le modèle à la fin de l'importation");

                InitVariables();

                OnCalque += Calque;
                OnRunAfterActivation += Rechercher_Fichier;
                OnRunOkCommand += RunOkCommand;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private CtrlTextBox _Texte_SelectionnerFichierInfos;
        private CtrlButton _Bouton_Parcourir;
        private CtrlTextListBox _TextListBox_AfficherInfos;
        private CtrlCheckBox _CheckBox_ComposantsExterne;
        private CtrlCheckBox _CheckBox_ToutReconstruire;

        String NomFichier = "";
        int Iteration = 0;

        private void InitVariables()
        {
            NomFichier = _pFichierInfos.GetValeur<String>();
            Iteration = 0;
        }

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Selectionnez le fichier : ");

                _Texte_SelectionnerFichierInfos = G.AjouterTexteBox();
                _Texte_SelectionnerFichierInfos.LectureSeule = true;

                _Bouton_Parcourir = G.AjouterBouton("Parcourir");
                _Bouton_Parcourir.OnButtonPress += delegate(Object Bouton)
                {
                    System.Windows.Forms.OpenFileDialog pDialogue = new System.Windows.Forms.OpenFileDialog();
                    pDialogue.Filter = "Fichier texte (*.txt)|*.txt|Tout les fichiers (*.*)|*.*";
                    pDialogue.Multiselect = false;
                    pDialogue.InitialDirectory = Path.GetDirectoryName(App.ModelDoc2.GetPathName());
                    pDialogue.RestoreDirectory = true;
                    
                    String pChemin = "";

                    if (pDialogue.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        pChemin = pDialogue.FileName;

                    PublierInfos(pChemin);
                };

                _TextListBox_AfficherInfos = G.AjouterTextListBox("Contenu du fichier :");
                _TextListBox_AfficherInfos.TouteHauteur = true;
                _TextListBox_AfficherInfos.Height = 90;
                _TextListBox_AfficherInfos.SelectionMultiple = true;

                G = _Calque.AjouterGroupe("Options");

                _CheckBox_ComposantsExterne = G.AjouterCheckBox(_pComposantsExterne);
                _CheckBox_ToutReconstruire = G.AjouterCheckBox(_pToutReconstruire);
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected String ScanFichier(String Dossier)
        {
            String pChemin = Path.Combine(Dossier, NomFichier);

            if (!File.Exists(pChemin))
            {
                try
                {
                    String DossierParent = Directory.GetParent(Dossier).FullName;
                    if (!String.IsNullOrWhiteSpace(DossierParent) && (Iteration++ < 5))
                        pChemin = ScanFichier(DossierParent);
                    else
                        pChemin = "";
                }
                catch
                {
                    pChemin = "";
                }
            }

            return pChemin;
        }

        protected void Rechercher_Fichier()
        {
            try
            {
                if (String.IsNullOrWhiteSpace(NomFichier))
                    return;

                // On regarde dans le repertoire du modèle courant
                String pChemin = ScanFichier(App.ModelDoc2.eDossier());

                PublierInfos(pChemin);
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private void PublierInfos(String Chemin)
        {
            try
            {
                List<String> pListe = new List<String>();

                if (String.IsNullOrWhiteSpace(Chemin) || !File.Exists(Chemin))
                    pListe.Add("");
                else
                    using (StreamReader sr = new StreamReader(Chemin, System.Text.Encoding.GetEncoding(1252)))
                    {
                        while (sr.Peek() > -1)
                            pListe.Add(sr.ReadLine());
                    }

                _Texte_SelectionnerFichierInfos.Text = Chemin;
                _TextListBox_AfficherInfos.Liste = pListe;
                _TextListBox_AfficherInfos.ToutSelectionner(false);
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RunOkCommand()
        {
            CmdImporterInfos Cmd = new CmdImporterInfos();

            Cmd.MdlBase = App.ModelDoc2;
            Cmd.ListeValeurs = _TextListBox_AfficherInfos.ListSelectedText.Count > 0 ? _TextListBox_AfficherInfos.ListSelectedText : _TextListBox_AfficherInfos.Liste;
            Cmd.ComposantsExterne = _CheckBox_ComposantsExterne.IsChecked;
            Cmd.ToutReconstruire = _CheckBox_ToutReconstruire.IsChecked;

            Cmd.Executer();
        }
    }
}
