using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;
using System.Collections.Generic;

namespace ModuleExportFichier
{
    namespace ModuleDxfDwg
    {
        [ModuleTypeDocContexte(eTypeDoc.Dessin),
            ModuleTitre("Exporter en Dxf/Dwg"),
            ModuleNom("ExportDxfDwg"),
            ModuleDescription("Exporter les feuilles en Dxf ou Dwg")
            ]
        public class PageDxfDwg : PageExportFichier
        {
            private Parametre VersionExport;
            private Parametre UtiliserPolicesACAD;
            private Parametre UtiliserStylesACAD;
            private Parametre FusionnerExtremites;
            private Parametre ToleranceFusion;

            private Parametre ExporterHauteQualite;
            private Parametre ExporterFeuilleEspacePapier;
            private Parametre ConvertirSplineToPolyligne;

            private Parametre CheminDernierDossier;


            public PageDxfDwg()
            {
                try
                {
                    FormatExport = _Config.AjouterParam("FormatExport", eTypeFichierExport.DWG, "Format");
                    VersionExport = _Config.AjouterParam("VersionExport", eDxfFormat.R2013, "Version");
                    UtiliserPolicesACAD = _Config.AjouterParam("UtiliserPoliceACAD", true, "Utiliser les polices AutoCAD");
                    UtiliserStylesACAD = _Config.AjouterParam("UtiliserStylesACAD", false, "Utiliser les styles AutoCAD");
                    FusionnerExtremites = _Config.AjouterParam("FusionnerExtremites", true, "Fusionner les extremités");
                    ToleranceFusion = _Config.AjouterParam("ToleranceFusion", 0.01, "Tolérance de fusion", "Tolérance de fusion");
                    ExporterHauteQualite = _Config.AjouterParam("ExporterHauteQualite", true, "Exporter en haute qualité");
                    ExporterFeuilleEspacePapier = _Config.AjouterParam("ExporterFeuilleEspacePapier", false, "Exporter les feuilles dans l'espace papier");
                    ConvertirSplineToPolyligne = _Config.AjouterParam("ConvertirSplineToPolyligne", false, "Convertir les splines en polylignes");

                    CheminDernierDossier = _Config.AjouterParam("CheminDernierDossier", "", "Dernier dossier utilisé");
                    if (CheminDernierDossier.GetValeur<String>().Contains(MdlBase.eDossier()))
                        MdlBase.eDrawingDoc().eSetDernierDossier(OutilsCommun.CheminRelatif(MdlBase.eDossier(), CheminDernierDossier.GetValeur<String>()));
                    else
                        MdlBase.eDrawingDoc().eSetDernierDossier("");

                    OnCalque += Calque;
                    OnRunOkCommand += RunOkCommand;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }
            private CtrlEnumComboBox<eTypeFichierExport, Intitule> _EnumComboBox_FormatExport;
            private CtrlEnumComboBox<eDxfFormat, Intitule> _EnumComboBox_VersionExport;
            private CtrlCheckBox _CheckBox_UtiliserPoliceACAD;
            private CtrlCheckBox _CheckBox_UtiliserStylesACAD;
            private CtrlCheckBox _CheckBox_FusionnerExtremites;
            private CtrlTextBox _TextBox_ToleranceFusion;
            private CtrlCheckBox _CheckBox_ExporterHauteQualite;
            private CtrlCheckBox _CheckBox_ExporterFeuilleEspacePapier;
            private CtrlCheckBox _CheckBox_ConvertirSplineToPolyligne;

            protected void Calque()
            {
                try
                {
                    G = _Calque.AjouterGroupe("Options");

                    _EnumComboBox_FormatExport = G.AjouterEnumComboBox<eTypeFichierExport, Intitule>(FormatExport);
                    _EnumComboBox_FormatExport.FiltrerEnum = eTypeFichierExport.DXF | eTypeFichierExport.DWG;

                    _EnumComboBox_VersionExport = G.AjouterEnumComboBox<eDxfFormat, Intitule>(VersionExport);

                    _CheckBox_UtiliserPoliceACAD = G.AjouterCheckBox(UtiliserPolicesACAD);
                    _CheckBox_UtiliserStylesACAD = G.AjouterCheckBox(UtiliserStylesACAD);

                    _CheckBox_FusionnerExtremites = G.AjouterCheckBox(FusionnerExtremites);
                    _TextBox_ToleranceFusion = G.AjouterTexteBox(ToleranceFusion);
                    _TextBox_ToleranceFusion.StdIndent();
                    _CheckBox_FusionnerExtremites.OnIsCheck += _TextBox_ToleranceFusion.IsEnable;
                    _CheckBox_ExporterHauteQualite = G.AjouterCheckBox(ExporterHauteQualite);
                    _CheckBox_ExporterHauteQualite.StdIndent();
                    _CheckBox_FusionnerExtremites.OnIsCheck += _CheckBox_ExporterHauteQualite.IsEnable;
                    _CheckBox_FusionnerExtremites.ApplyParam();

                    _CheckBox_ConvertirSplineToPolyligne = G.AjouterCheckBox(ConvertirSplineToPolyligne);

                    _CheckBox_ExporterFeuilleEspacePapier = G.AjouterCheckBox(ExporterFeuilleEspacePapier);

                    AvecIndice = false;
                    AjouterCalqueDossier();

                    _EnumComboBox_FormatExport.OnSelectionChanged += delegate (Object sender, int Item)
                    {
                        _SelectionnerDossier.TypeFichier = _EnumComboBox_FormatExport.Val;
                        _SelectionnerDossier.Maj();
                        _DernierDossier.TypeFichier = _EnumComboBox_FormatExport.Val;
                        _DernierDossier.Maj();
                    };

                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private Boolean Svg_DxfDwg_ExporterFeuilleDansEspacePapier = false;

            private void AppliquerOptions()
            {
                Sw.DxfDwg_Version = _EnumComboBox_VersionExport.Val;
                Sw.DxfDwg_PolicesAutoCAD = _CheckBox_UtiliserPoliceACAD.IsChecked;
                Sw.DxfDwg_StylesAutoCAD = _CheckBox_UtiliserStylesACAD.IsChecked;
                Sw.DxfDwg_JoindreExtremites = _CheckBox_FusionnerExtremites.IsChecked;
                Sw.DxfDwg_JoindreExtremitesTolerance = _TextBox_ToleranceFusion.Text.eToDouble();
                Sw.DxfDwg_JoindreExtremitesHauteQualite = _CheckBox_ExporterHauteQualite.IsChecked;
                Sw.DxfDwg_ExporterSplineEnPolyligne = _CheckBox_ConvertirSplineToPolyligne.IsChecked;

                Svg_DxfDwg_ExporterFeuilleDansEspacePapier = Sw.DxfDwg_ExporterFeuilleDansEspacePapier;
                Sw.DxfDwg_ExporterFeuilleDansEspacePapier = _CheckBox_ExporterFeuilleEspacePapier.IsChecked;
            }



            protected void RunOkCommand()
            {
                AppliquerOptions();

                CmdDxfDwg Cmd = new CmdDxfDwg();
                Cmd.Dessin = MdlBase.eDrawingDoc();
                Cmd.typeExport = _EnumComboBox_FormatExport.Val;
                Cmd.CheminDossier = NomDossier;
                Cmd.Feuille = MdlBase.eDrawingDoc().eFeuilleActive();
                Cmd.ToutesLesFeuilles = _CheckBox_ToutesLesFeuilles.IsChecked;
                Cmd.NomFichier = NomFichierComplet;

                CheminDernierDossier.SetValeur<String>(NomDossier);

                Cmd.Executer();

                RetablirOptions();
            }

            private void RetablirOptions()
            {
                Sw.DxfDwg_ExporterFeuilleDansEspacePapier = Svg_DxfDwg_ExporterFeuilleDansEspacePapier;
                ExporterFeuilleEspacePapier.SetValeur(Svg_DxfDwg_ExporterFeuilleDansEspacePapier);
            }
        }
    }
}
