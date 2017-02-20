using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.IO;

namespace ModuleExportFichier
{
    namespace ModulePdf
    {
        [ModuleTypeDocContexte(eTypeDoc.Dessin),
            ModuleTitre("Exporter en Pdf"),
            ModuleNom("ExportPdf"),
            ModuleDescription("Exporter les feuilles en Pdf")
            ]
        public class PagePdf : PageExportFichier
        {
            private Parametre ExporterEnCouleur;
            private Parametre IncorporerLesPolices;
            private Parametre ExporterEnHauteQualite;
            private Parametre ImprimerEnTeteEtPiedDePage;
            private Parametre EpaisseursDeLigneDeImprimante;


            public PagePdf()
            {
                try
                {
                    FormatExport = _Config.AjouterParam("FormatExport", eTypeFichierExport.PDF, "Format");
                    ExporterEnCouleur = _Config.AjouterParam("ExporterEnCouleur", true, "Exporter en couleur");
                    IncorporerLesPolices = _Config.AjouterParam("IncorporerLesPolices", true, "Incorporer les polices");
                    ExporterEnHauteQualite = _Config.AjouterParam("ExporterEnHauteQualite", true, "Exporter en haute qualité");
                    ImprimerEnTeteEtPiedDePage = _Config.AjouterParam("ImprimerEnTeteEtPiedDePage", false, "Imprimer l'entête et le pied de page");
                    EpaisseursDeLigneDeImprimante = _Config.AjouterParam("EpaisseursDeLigneDeImprimante", false, "Epaisseur de lignes de l'imprimante");

                    OnCalque += Calque;
                    OnRunOkCommand += RunOkCommand;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private CtrlCheckBox _CheckBox_ExporterEnCouleur;
            private CtrlCheckBox _CheckBox_IncorporerLesPolices;
            private CtrlCheckBox _CheckBox_ExporterEnHauteQualite;
            private CtrlCheckBox _CheckBox_ImprimerEnTeteEtPiedDePage;
            private CtrlCheckBox _CheckBox_EpaisseursDeLigneDeImprimante;

            protected void Calque()
            {
                try
                {
                    G = _Calque.AjouterGroupe("Options");

                    _CheckBox_ExporterEnCouleur = G.AjouterCheckBox(ExporterEnCouleur);
                    _CheckBox_IncorporerLesPolices = G.AjouterCheckBox(IncorporerLesPolices);
                    _CheckBox_ExporterEnHauteQualite = G.AjouterCheckBox(ExporterEnHauteQualite);
                    _CheckBox_ImprimerEnTeteEtPiedDePage = G.AjouterCheckBox(ImprimerEnTeteEtPiedDePage);
                    _CheckBox_EpaisseursDeLigneDeImprimante = G.AjouterCheckBox(EpaisseursDeLigneDeImprimante);

                    AjouterCalqueDossier();
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private void AppliquerOptions()
            {
                Sw.Pdf_ExporterEnCouleur = _CheckBox_ExporterEnCouleur.IsChecked;
                Sw.Pdf_IncorporerLesPolices = _CheckBox_IncorporerLesPolices.IsChecked;
                Sw.Pdf_ExporterEnHauteQualite = _CheckBox_ExporterEnHauteQualite.IsChecked;
                Sw.Pdf_ImprimerEnTeteEtPiedDePage = _CheckBox_ImprimerEnTeteEtPiedDePage.IsChecked;
                Sw.Pdf_UtiliserLesEpaisseursDeLigneDeImprimante = _CheckBox_EpaisseursDeLigneDeImprimante.IsChecked;
            }

            protected void RunOkCommand()
            {
                AppliquerOptions();

                CmdPdf Cmd = new CmdPdf();
                Cmd.Dessin = App.DrawingDoc;
                Cmd.typeExport = eTypeFichierExport.PDF;
                Cmd.CheminDossier = NomDossier;
                Cmd.ToutesLesFeuilles = _CheckBox_ToutesLesFeuilles.IsChecked;
                Cmd.Feuille = (Sheet)App.DrawingDoc.GetCurrentSheet();
                Cmd.NomFichier = NomFichierComplet;

                Cmd.Executer();
            }
        }
    }
}
