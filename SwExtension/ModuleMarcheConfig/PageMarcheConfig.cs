using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ModuleMarcheConfig
{
    public abstract class PageMarcheConfig : BoutonPMPManager
    {
        protected CtrlCheckBox _CheckBox_EnregistrerSelection;

        protected Boolean FiltreFace(CtrlSelectionBox SelBox, Object Cp, int selType, Parametre param)
        {
            if (selType == (int)swSelectType_e.swSelFACES)
                return true;

            if (Cp.IsRef())
                SelectFace(SelBox, Cp as Component2, param);

            return false;
        }

        protected Boolean SelectFace(CtrlSelectionBox SelBox, Component2 Cp, Parametre param)
        {
            if (Cp.IsNull()) return false;

            String cFace = param.GetValeur<String>();
            Face2 Face = Cp.eChercherFace(cFace);
            if (Face.IsRef())
                MdlBase.eSelectMulti(Face, SelBox.Marque, true);

            return true;
        }

        protected Boolean FiltreEsquisse(CtrlSelectionBox SelBox, Object selection, int selType, Parametre param)
        {
            if (selType == (int)swSelectType_e.swSelSKETCHES)
                return true;

            Component2 Cp = selection as Component2;
            if (Cp.IsRef() && !SelectEsquisse(SelBox, Cp, param))
            {
                foreach (var C in Cp.eRecListeComposant())
                    if (SelectEsquisse(SelBox, C, param))
                        break;
            }

            return false;
        }

        protected Boolean SelectEsquisse(CtrlSelectionBox SelBox, Component2 Cp, Parametre param)
        {
            if (Cp.IsNull()) return false;

            String cEsquisseConfig = param.GetValeur<String>();
            Feature Esquisse = Cp.eChercherFonction(f => { return Regex.IsMatch(f.Name, cEsquisseConfig); }, true);
            if (Esquisse.IsRef())
                MdlBase.eSelectMulti(Esquisse, SelBox.Marque, true);

            return true;
        }

        protected Boolean FiltrePlan(CtrlSelectionBox SelBox, Object selection, int selType, Parametre param)
        {
            if (selType == (int)swSelectType_e.swSelDATUMPLANES)
                return true;

            Component2 Cp = selection as Component2;
            if (Cp.IsRef())
            {
                List<Component2> Liste = Cp.eListeComposantParent();
                Liste.Insert(0, Cp);
                SelectPlan(SelBox, Liste.Last(), param);
            }

            return false;
        }

        protected Boolean SelectPlan(CtrlSelectionBox SelBox, Component2 Cp, Parametre param)
        {
            if (Cp.IsNull()) return false;

            String cPlanContrainte = param.GetValeur<String>();
            Feature Plan = Cp.eChercherFonction(f => { return Regex.IsMatch(f.Name, cPlanContrainte); }, false);

            if (Plan.IsRef())
                MdlBase.eSelectMulti(Plan, SelBox.Marque, true);

            return false;
        }

        protected void SvgNomComposant(Object SelBox, Parametre Param)
        {
            Component2 Cp = MdlBase.eSelect_RecupererComposant(1, ((CtrlSelectionBox)SelBox).Marque);
            if (Cp.IsRef() && _CheckBox_EnregistrerSelection.IsChecked)
                Param.SetValeur(Cp.eNomSansExt());
        }

        protected void SvgNomFonction(Object SelBox, Parametre Param)
        {
            Feature F = MdlBase.eSelect_RecupererObjet<Feature>(1, ((CtrlSelectionBox)SelBox).Marque);
            if (F.IsRef() && _CheckBox_EnregistrerSelection.IsChecked)
                Param.SetValeur(F.Name);
        }
    }
}
