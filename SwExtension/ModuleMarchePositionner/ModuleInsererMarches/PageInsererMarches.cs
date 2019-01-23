using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ModuleMarchePositionner
{
    namespace ModuleInsererMarches
    {
        [ModuleTypeDocContexte(eTypeDoc.Assemblage),
            ModuleTitre("Inserer les marches"),
            ModuleNom("InsererMarches"),
            ModuleDescription("Inserer les marches suivant une fonction de répétition." +
                                "\r\nLa marche doit être contrainte au corps de base de la répétition"
            )
            ]
        public class PageInsererMarches : PageMarchePositionner
        {
            private Parametre _pMarche;
            private Parametre _pPieceRepet;
            private Parametre _pFonctionRepet;

            public PageInsererMarches()
            {
                try
                {
                    _pMarche = _Config.AjouterParam("Marche", "AM", "Selectionnez la marche");
                    _pPieceRepet = _Config.AjouterParam("PieceRepet", "PP01", "Selectionnez la platine gauche");
                    _pFonctionRepet = _Config.AjouterParam("FonctionRepet", "F_Dessus", "Selectionnez la face du dessus");

                    OnCalque += Calque;
                    OnRunOkCommand += RunOkCommand;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private CtrlSelectionBox _Select_Marche;
            private CtrlSelectionBox _Select_FonctionRepet;
            private CtrlSelectionBox _Select_Contraintes_PointCorps;
            private CtrlSelectionBox _Select_Contraintes_PointMarche;
            private CtrlSelectionBox _Select_Contraintes_ArreteCorps;
            private CtrlSelectionBox _Select_Contraintes_AxeMarche;
            private CtrlSelectionBox _Select_Contraintes_PlanComp;
            private CtrlSelectionBox _Select_Contraintes_PlanMarche;

            private CtrlButton _Button_Preselection;

            protected void Calque()
            {
                try
                {
                    Groupe G;
                    G = _Calque.AjouterGroupe("Appliquer");

                    _CheckBox_EnregistrerSelection = G.AjouterCheckBox("Enregistrer les selections");
                    _Button_Preselection = G.AjouterBouton("Preselectionner");
                    _Button_Preselection.OnButtonPress += delegate (object sender) { PreSelection(); };

                    G = _Calque.AjouterGroupe("Marche" + " ( " + _pMarche.GetValeur<String>() + " )");

                    _Select_Marche = G.AjouterSelectionBox("", "Selectionnez la marche");
                    _Select_Marche.SelectionMultipleMemeEntite = false;
                    _Select_Marche.SelectionDansMultipleBox = false;
                    _Select_Marche.UneSeuleEntite = true;
                    _Select_Marche.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
                    _Select_Marche.OnSubmitSelection += FiltreMarche;

                    _Select_Marche.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, _pMarche); };
                    _Select_Marche.OnSelectionChanged += EffacerContraintes;

                    G = _Calque.AjouterGroupe("Fonction de répetition");

                    _Select_FonctionRepet = G.AjouterSelectionBox("Fonction" + " ( " + _pFonctionRepet.GetValeur<String>() + "@" + _pPieceRepet.GetValeur<String>() + " )", "Selectionnez le plan");
                    _Select_FonctionRepet.SelectionMultipleMemeEntite = false;
                    _Select_FonctionRepet.SelectionDansMultipleBox = false;
                    _Select_FonctionRepet.UneSeuleEntite = true;
                    _Select_FonctionRepet.FiltreSelection(new List<swSelectType_e>() {
                        swSelectType_e.swSelBODYFEATURES,
                        swSelectType_e.swSelFACES
                    });
                    _Select_FonctionRepet.OnSubmitSelection += FiltreFonctionRepetition;
                    _Select_FonctionRepet.OnSelectionChanged += EffacerContraintes;

                    // Svg des parametres
                    _Select_FonctionRepet.OnApplyOnSelection += delegate (Object Box) { SvgNomComposant(Box, _pPieceRepet); };
                    _Select_FonctionRepet.OnApplyOnSelection += delegate (Object Box) { SvgNomFonction(Box, _pFonctionRepet); };

                    _Select_Marche.OnApplyOnSelection += _Select_FonctionRepet.GainedFocus;

                    G = _Calque.AjouterGroupe("Contrainte point");

                    _Select_Contraintes_PointCorps = G.AjouterSelectionBox("Point sur le corps");
                    _Select_Contraintes_PointCorps.SelectionMultipleMemeEntite = false;
                    _Select_Contraintes_PointCorps.SelectionDansMultipleBox = false;
                    _Select_Contraintes_PointCorps.UneSeuleEntite = true;
                    _Select_Contraintes_PointCorps.FiltreSelection(swSelectType_e.swSelVERTICES);

                    _Select_Contraintes_PointMarche = G.AjouterSelectionBox("Point dans la marche");
                    _Select_Contraintes_PointMarche.SelectionMultipleMemeEntite = false;
                    _Select_Contraintes_PointMarche.SelectionDansMultipleBox = false;
                    _Select_Contraintes_PointMarche.UneSeuleEntite = true;
                    _Select_Contraintes_PointMarche.FiltreSelection(swSelectType_e.swSelEXTSKETCHPOINTS);

                    G = _Calque.AjouterGroupe("Contrainte axe");

                    _Select_Contraintes_ArreteCorps = G.AjouterSelectionBox("Arrete sur le corps");
                    _Select_Contraintes_ArreteCorps.SelectionMultipleMemeEntite = false;
                    _Select_Contraintes_ArreteCorps.SelectionDansMultipleBox = false;
                    _Select_Contraintes_ArreteCorps.UneSeuleEntite = true;
                    _Select_Contraintes_ArreteCorps.FiltreSelection(swSelectType_e.swSelEDGES);

                    _Select_Contraintes_AxeMarche = G.AjouterSelectionBox("Axe dans la marche");
                    _Select_Contraintes_AxeMarche.SelectionMultipleMemeEntite = false;
                    _Select_Contraintes_AxeMarche.SelectionDansMultipleBox = false;
                    _Select_Contraintes_AxeMarche.UneSeuleEntite = true;
                    _Select_Contraintes_AxeMarche.FiltreSelection(swSelectType_e.swSelDATUMAXES);

                    G = _Calque.AjouterGroupe("Contrainte plan");

                    _Select_Contraintes_PlanComp = G.AjouterSelectionBox("Plan dans le composant");
                    _Select_Contraintes_PlanComp.SelectionMultipleMemeEntite = false;
                    _Select_Contraintes_PlanComp.SelectionDansMultipleBox = false;
                    _Select_Contraintes_PlanComp.UneSeuleEntite = true;
                    _Select_Contraintes_PlanComp.FiltreSelection(swSelectType_e.swSelDATUMPLANES);

                    _Select_Contraintes_PlanMarche = G.AjouterSelectionBox("Plan dans la marche");
                    _Select_Contraintes_PlanMarche.SelectionMultipleMemeEntite = false;
                    _Select_Contraintes_PlanMarche.SelectionDansMultipleBox = false;
                    _Select_Contraintes_PlanMarche.UneSeuleEntite = true;
                    _Select_Contraintes_PlanMarche.FiltreSelection(swSelectType_e.swSelDATUMPLANES);

                    // OnCheck, on enregistre les parametres
                    _CheckBox_EnregistrerSelection.OnCheck += _Select_Marche.ApplyOnSelection;
                    _CheckBox_EnregistrerSelection.OnCheck += _Select_FonctionRepet.ApplyOnSelection;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private Entity PointCorps = null;
            private SketchPoint PointMarche = null;
            private Entity ArreteCorps = null;
            private Feature AxeMarche = null;
            private Feature PlanComp = null;
            private Feature PlanMarche = null;

            private Boolean ContraintesVide = true;

            public void EffacerContraintes(Object box = null, int nb = 0)
            {
                try
                {
                    if (ContraintesVide || (box.IsRef() && (nb > 0))) return;

                    PointCorps = null;
                    ArreteCorps = null;
                    PlanComp = null;
                    AxeMarche = null;
                    PlanComp = null;
                    PlanMarche = null;

                    if (_Select_Contraintes_PointCorps.Nb > 0)
                    {
                        List<Entity> le = MdlBase.eSelect_RecupererListeObjets<Entity>(_Select_Contraintes_PointCorps.Marque);

                        foreach (Entity e in le)
                            e.DeSelect();
                    }

                    if (_Select_Contraintes_PointMarche.Nb > 0)
                    {
                        List<SketchPoint> le = MdlBase.eSelect_RecupererListeObjets<SketchPoint>(_Select_Contraintes_PointMarche.Marque);

                        foreach (SketchPoint f in le)
                            f.DeSelect();
                    }

                    if (_Select_Contraintes_ArreteCorps.Nb > 0)
                    {
                        List<Entity> lm = MdlBase.eSelect_RecupererListeObjets<Entity>(_Select_Contraintes_ArreteCorps.Marque);
                        foreach (Entity e in lm)
                            e.DeSelect();
                    }

                    if (_Select_Contraintes_AxeMarche.Nb > 0)
                    {
                        List<Feature> lm = MdlBase.eSelect_RecupererListeObjets<Feature>(_Select_Contraintes_AxeMarche.Marque);
                        foreach (Feature f in lm)
                            f.DeSelect();
                    }

                    if (_Select_Contraintes_PlanComp.Nb > 0)
                    {
                        List<Feature> lm = MdlBase.eSelect_RecupererListeObjets<Feature>(_Select_Contraintes_PlanComp.Marque);
                        foreach (Feature f in lm)
                            f.DeSelect();
                    }

                    if (_Select_Contraintes_PlanMarche.Nb > 0)
                    {
                        List<Feature> lm = MdlBase.eSelect_RecupererListeObjets<Feature>(_Select_Contraintes_PlanMarche.Marque);
                        foreach (Feature f in lm)
                            f.DeSelect();
                    }
                    ContraintesVide = true;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            public void SelectContraintes()
            {
                try
                {
                    EffacerContraintes();

                    Component2 cpMarche = MdlBase.eSelect_RecupererComposant(1, _Select_Marche.Marque);
                    Feature fRepet = MdlBase.eSelect_RecupererObjet<Feature>(1, _Select_FonctionRepet.Marque);
                    Component2 cpRepet = MdlBase.eSelect_RecupererComposant(1, _Select_FonctionRepet.Marque);
                    Object[] Contraintes = cpMarche.GetMates();

                    if (cpMarche.IsNull() || fRepet.IsNull() || cpRepet.IsNull() || Contraintes.IsNull()) return;

                    CurveDrivenPatternFeatureData def = fRepet.GetDefinition();
                    Object[] tabB = def.PatternBodyArray;
                    Body2 b = (Body2)tabB[0];

                    foreach (Mate2 m in Contraintes)
                    {
                        MateEntity2 m1 = m.MateEntity(0);
                        MateEntity2 m2 = m.MateEntity(1);

                        if (AppartientAuCorps(cpRepet, b, m1))
                        {
                            if ((m1.ReferenceType2 == (int)swSelectType_e.swSelVERTICES) && PointCorps.IsNull())
                            {
                                PointCorps = m1.Reference;
                                PointMarche = m2.Reference;
                            }
                            else if ((m1.ReferenceType2 == (int)swSelectType_e.swSelEDGES) && ArreteCorps.IsNull())
                            {
                                ArreteCorps = m1.Reference;
                                AxeMarche = m2.Reference;
                            }
                        }
                        else if (AppartientAuComposant(cpRepet, m1) && (m1.ReferenceType2 == (int)swSelectType_e.swSelDATUMPLANES) && PlanComp.IsNull())
                        {
                            PlanComp = m1.Reference;
                            PlanMarche = m2.Reference;
                        }
                        else if (AppartientAuCorps(cpRepet, b, m2))
                        {
                            if ((m2.ReferenceType2 == (int)swSelectType_e.swSelVERTICES) && PointCorps.IsNull())
                            {
                                PointCorps = m2.Reference;
                                PointMarche = m1.Reference;
                            }
                            else if ((m2.ReferenceType2 == (int)swSelectType_e.swSelEDGES) && ArreteCorps.IsNull())
                            {
                                ArreteCorps = m2.Reference;
                                AxeMarche = m1.Reference;
                            }
                        }
                        else if (AppartientAuComposant(cpRepet, m2) && (m2.ReferenceType2 == (int)swSelectType_e.swSelDATUMPLANES) && PlanComp.IsNull())
                        {
                            PlanComp = m2.Reference;
                            PlanMarche = m1.Reference;
                        }
                    }

                    MdlBase.eSelectMulti(PointCorps, _Select_Contraintes_PointCorps.Marque, true);
                    MdlBase.eSelectMulti(PointMarche, _Select_Contraintes_PointMarche.Marque, true);
                    MdlBase.eSelectMulti(ArreteCorps, _Select_Contraintes_ArreteCorps.Marque, true);
                    MdlBase.eSelectMulti(AxeMarche, _Select_Contraintes_AxeMarche.Marque, true);
                    MdlBase.eSelectMulti(PlanComp, _Select_Contraintes_PlanComp.Marque, true);
                    MdlBase.eSelectMulti(PlanMarche, _Select_Contraintes_PlanMarche.Marque, true);

                    ContraintesVide = false;
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            private Boolean AppartientAuComposant(Component2 cp, MateEntity2 m)
            {
                if (cp.Name2 == m.ReferenceComponent.Name2)
                    return true;

                return false;
            }

            private Boolean AppartientAuCorps(Component2 cp, Body2 corps, MateEntity2 m)
            {
                if (cp.Name2 != m.ReferenceComponent.Name2) return false;

                Edge e;

                if (m.ReferenceType2 == (int)swSelectType_e.swSelVERTICES)
                {
                    Vertex v = m.Reference;
                    e = v.GetEdges()[0];
                }
                else if (m.ReferenceType2 == (int)swSelectType_e.swSelEDGES)
                {
                    e = m.Reference;
                }
                else if (m.ReferenceType2 == (int)swSelectType_e.swSelFACES)
                {
                    e = m.Reference;
                }
                else
                    return false;

                Body2 b = e.GetBody();

                if (b.Name == corps.Name)
                    return true;

                return false;
            }

            public Boolean FiltreMarche(Object SelBox, Object selection, int selType, String itemText)
            {
                Boolean result = SelectionnerComposant1erNvx(SelBox, selection, selType, itemText);

                SelectContraintes();

                return result;
            }

            public Boolean FiltreFonctionRepetition(Object SelBox, Object selection, int selType, String itemText)
            {
                EffacerContraintes();

                if (selType == (int)swSelectType_e.swSelBODYFEATURES)
                {
                    Feature f = selection as Feature;
                    if (f.GetTypeName2() == FeatureType.swTnCurvePattern)
                    {
                        SelectContraintes();
                        return true;
                    }
                }
                else
                {
                    try
                    {
                        Face2 face = selection as Face2;
                        Body2 b = face.GetBody();
                        Entity e = face as Entity;
                        Component2 c = e.GetComponent();

                        String cNomFonc = _pFonctionRepet.GetValeur<String>();
                        Feature F = b.eChercherFonction(f => { return Regex.IsMatch(f.Name, cNomFonc); }, false);

                        if (F.IsNull())
                            F = b.eChercherFonction(f => { return f.GetTypeName2() == FeatureType.swTnCurvePattern; }, false);

                        F = c.FeatureByName(F.Name);

                        SelectFonctionRepetition((CtrlSelectionBox)SelBox, F);
                    }
                    catch (Exception e)
                    { this.LogMethode(new Object[] { e }); }
                }

                return false;
            }

            public Boolean SelectFonctionRepetition(CtrlSelectionBox SelBox, Feature F)
            {
                if (F.IsRef())
                {
                    List<Feature> lcp = MdlBase.eSelect_RecupererListeObjets<Feature>(SelBox.Marque);
                    if (lcp.Count > 0)
                    {
                        foreach (Feature f in lcp)
                            f.DeSelect();
                    }
                    else
                    {
                        MdlBase.eSelectMulti(F, SelBox.Marque, true);
                        SelectContraintes();
                    }
                }

                return false;
            }

            protected void PreSelection()
            {
                try
                {
                    MdlBase.ClearSelection2(true);

                    Component2 Marche = MdlBase.eRecChercherComposant(c =>
                    {
                        return Regex.IsMatch(c.Name2, _pMarche.GetValeur<String>())
                     && !c.IsSuppressed();
                    },
                    c => { return false; }
                    );

                    if (Marche.IsRef())
                        MdlBase.eSelectMulti(Marche, _Select_Marche.Marque);

                    Component2 PieceRepet = MdlBase.eRecChercherComposant(c =>
                    {
                        return Regex.IsMatch(c.Name2, _pPieceRepet.GetValeur<String>())
                       && !c.IsSuppressed();
                    });

                    if (PieceRepet.IsRef())
                    {
                        Feature F = PieceRepet.eChercherFonction(f => { return Regex.IsMatch(f.Name, _pFonctionRepet.GetValeur<String>()); }, false);
                        if (F.IsRef())
                            SelectFonctionRepetition(_Select_FonctionRepet, F);
                    }
                }
                catch (Exception e)
                { this.LogMethode(new Object[] { e }); }
            }

            protected void RunOkCommand()
            {
                CmdInsererMarches Cmd = new CmdInsererMarches();
                Cmd.MdlBase = MdlBase;
                Cmd.Marche = MdlBase.eSelect_RecupererComposant(1, _Select_Marche.Marque);
                Cmd.ComposantRepetition = MdlBase.eSelect_RecupererComposant(1, _Select_FonctionRepet.Marque);
                Cmd.FonctionRepetition = MdlBase.eSelect_RecupererObjet<Feature>(1, _Select_FonctionRepet.Marque);
                Cmd.Point = PointCorps;
                Cmd.Arrete = ArreteCorps;
                Cmd.Plan = PlanComp;
                Cmd.PointMarche = PointMarche;
                Cmd.AxeMarche = AxeMarche;
                Cmd.PlanMarche = PlanMarche;

                MdlBase.ClearSelection2(true);

                Cmd.Executer();
            }
        }
    }
}
