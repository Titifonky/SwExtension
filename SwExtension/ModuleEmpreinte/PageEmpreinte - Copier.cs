using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ModuleEmpreinte
{
    //[ModuleTypeDocContexte(eTypeDoc.Assemblage),
    //    ModuleTitre("Empreinte entre les composants"),
    //    ModuleNom("Empreinte"),
    //    ModuleDescription("Empreinte entre les composants sélectionnés")
    //    ]
    //public class PageEmpreinteOld : BoutonPMPManager
    //{
    //    private Parametre PrefixeBase;
    //    private Parametre PrefixeEmpreinte;
    //    private Parametre PropEmpreinte;
    //    private Parametre PropPrefixeEmpreinte;
    //    private Parametre PropTarauderEmpreinte;

    //    public PageEmpreinteOld()
    //    {
    //        PrefixeBase = _Config.AjouterParam("PrefixeBase", "", "Filtrer les prefixes (?*#):");
    //        PrefixeEmpreinte = _Config.AjouterParam("PrefixeEmpreinte", "", "Filtrer les prefixes (?*#):");
    //        PropEmpreinte = _Config.AjouterParam("PropEmpreinte", "Empreinte", "Propriete \"Empreinte\"");
    //        PropPrefixeEmpreinte = _Config.AjouterParam("PropPrefixeEmpreinte", "PrefixeEmpreinte", "Propriete \"PrefixeEmpreinte\"");
    //        PropTarauderEmpreinte = _Config.AjouterParam("PropTarauderEmpreinte", "TarauderEmpreinte", "Propriete \"TarauderEmpreinte\"");

    //        Com.NomPropEmpreinte = PropEmpreinte.GetValeur<String>();
    //        Com.NomPropPrefixe = PropPrefixeEmpreinte.GetValeur<String>();

    //        OnCalque += Calque;
    //        OnRunOkCommand += RunOkCommand;
    //        OnRunAfterClose += RunAfterClose;
    //    }

    //    private CtrlSelectionBox _Select_CompBase;
    //    private CtrlSelectionBox _Select_CompEmpreinte;
    //    private CtrlTextBox _TextBox_PrefixeBase;
    //    private CtrlTextBox _TextBox_PrefixeEmpreinte;

    //    private CtrlButton _Button_PrefixeBase;
    //    private CtrlButton _Button_PrefixeEmpreinte;
    //    private CtrlCheckBox _CheckBox_MasquerLesEmpreintes;

    //    protected void Calque()
    //    {
    //        try
    //        {
    //            Groupe G;

    //            G = _Calque.AjouterGroupe("Selectionner les composants de base");

    //            _Select_CompBase = G.AjouterSelectionBox("", "Selectionnez les composants");
    //            _Select_CompBase.SelectionMultipleMemeEntite = false;
    //            _Select_CompBase.SelectionDansMultipleBox = false;
    //            _Select_CompBase.UneSeuleEntite = false;
    //            _Select_CompBase.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
    //            _Select_CompBase.OnSubmitSelection += SelectionnerPiece;
    //            _Select_CompBase.Hauteur = 8;
    //            _Select_CompBase.Focus = true;

    //            _TextBox_PrefixeBase = G.AjouterTexteBox(PrefixeBase, true);
    //            _Button_PrefixeBase = G.AjouterBouton("Rechercher composants");
    //            _Button_PrefixeBase.OnButtonPress += delegate (Object sender) { RechercherComp(_Select_CompBase, _TextBox_PrefixeBase.Text); };

    //            G = _Calque.AjouterGroupe("Selectionner les composants empreinte");

    //            _Select_CompEmpreinte = G.AjouterSelectionBox("", "Selectionnez les composants");
    //            _Select_CompEmpreinte.SelectionMultipleMemeEntite = false;
    //            _Select_CompEmpreinte.SelectionDansMultipleBox = false;
    //            _Select_CompEmpreinte.UneSeuleEntite = false;
    //            _Select_CompEmpreinte.FiltreSelection(swSelectType_e.swSelCOMPONENTS, swSelectType_e.swSelFACES);
    //            _Select_CompEmpreinte.OnSubmitSelection += SelectionnerPiece;
    //            _Select_CompEmpreinte.Hauteur = 8;

    //            _TextBox_PrefixeEmpreinte = G.AjouterTexteBox(PrefixeEmpreinte, true);
    //            _Button_PrefixeEmpreinte = G.AjouterBouton("Rechercher empreintes");
    //            _Button_PrefixeEmpreinte.OnButtonPress += delegate (Object sender) { RechercherComp(_Select_CompEmpreinte, _TextBox_PrefixeEmpreinte.Text); };

    //            _CheckBox_MasquerLesEmpreintes = G.AjouterCheckBox("Masquer toutes les empreintes");

    //        }
    //        catch (Exception e)
    //        { this.LogMethode(new Object[] { e }); }
    //    }

    //    private Dictionary<Component2, int> lstComps = new Dictionary<Component2, int>();

    //    private void RechercherComp(CtrlSelectionBox box, String pattern)
    //    {
    //        try
    //        {
    //            if (String.IsNullOrWhiteSpace(pattern))
    //                return;

    //            String[] listePattern = pattern.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

    //            ModelDoc2 mdl = App.ModelDoc2;

    //            List<Component2> lcp = mdl.eSelect_RecupererListeObjets<Component2>(box.Marque);
    //            foreach (Component2 c in lcp)
    //                c.eDeSelectById(mdl);

    //            box.Focus = true;

    //            RetablirVisibiliteComposants();

    //            mdl.eRecParcourirComposants(
    //                c =>
    //                {

    //                    if (!c.IsSuppressed() && (c.TypeDoc() == eTypeDoc.Piece))
    //                    {
    //                        if (c.ePropExiste(Com.NomPropEmpreinte) && (c.eProp(Com.NomPropEmpreinte) == "1"))
    //                        {
    //                            if (TestStringLikeListePattern(c.eProp(Com.NomPropPrefixe), listePattern))
    //                            {
    //                                lstComps.Add(c, c.Visible);

    //                                if (c.Visible == (int)swComponentVisibilityState_e.swComponentHidden)
    //                                    c.Visible = (int)swComponentVisibilityState_e.swComponentVisible;
    //                            }
    //                        }
    //                    }
    //                    return false;
    //                },
    //                null
    //                );

    //            mdl.eSelectMulti(lstComps.Keys.ToList(), box.Marque, true);
    //        }
    //        catch (Exception e)
    //        { this.LogMethode(new Object[] { e }); }

    //    }

    //    private Boolean TestStringLikeListePattern(String s, String[] Liste)
    //    {
    //        foreach (var t in Liste)
    //        {
    //            if (s.eIsLike(t))
    //                return true;
    //        }

    //        return false;
    //    }

    //    protected void RunOkCommand()
    //    {

    //        List<Component2> ListeCompBase = App.ModelDoc2.eSelect_RecupererListeObjets<Component2>(_Select_CompBase.Marque);
    //        List<Component2> ListeCompEmpreinte = App.ModelDoc2.eSelect_RecupererListeObjets<Component2>(_Select_CompEmpreinte.Marque);

            

    //        CmdEmpreinte Cmd = new CmdEmpreinte();
    //        Cmd.MdlBase = App.ModelDoc2;
    //        Cmd.ListeCompBase = ListeCompBase;
    //        Cmd.ListeCompEmpreinte = ListeCompEmpreinte;

    //        Cmd.Executer();
    //    }

    //    protected void RetablirVisibiliteComposants()
    //    {
    //        foreach (var cp in lstComps.Keys)
    //        {
    //            if (lstComps[cp] == (int)swComponentVisibilityState_e.swComponentHidden)
    //                cp.Visible = lstComps[cp];
    //        }

    //        lstComps.Clear();
    //    }

    //    protected void RunAfterClose()
    //    {
    //        List<Component2> ListeCompEmpreinte = App.ModelDoc2.eSelect_RecupererListeObjets<Component2>(_Select_CompEmpreinte.Marque);

    //        if (_CheckBox_MasquerLesEmpreintes.IsChecked == true)
    //        {
    //            WindowLog.Ecrire("Masque les composants");
    //            foreach (Component2 c in ListeCompEmpreinte)
    //                c.Visible = (int)swComponentVisibilityState_e.swComponentHidden;

    //            lstComps.Clear();
    //        }
    //        else
    //            RetablirVisibiliteComposants();
    //    }
    //}

    //public static class Com
    //{
    //    public static String NomPropEmpreinte;
    //    public static String NomPropPrefixe;

    //    public const String NOM_EMPREINTE = "Empreinte";

    //    public static String ValPropEmpreinte(this Component2 cp)
    //    {
    //        if (cp.ePropExiste(NomPropEmpreinte) && (cp.eProp(NomPropEmpreinte) == "1"))
    //        {
    //            return cp.eProp(NomPropPrefixe);
    //        }

    //        return "";
    //    }
    //}
}
