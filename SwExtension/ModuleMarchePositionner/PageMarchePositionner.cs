using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;

namespace ModuleMarchePositionner
{
    public abstract class PageMarchePositionner : BoutonPMPManager
    {
        protected CtrlCheckBox _CheckBox_EnregistrerSelection;

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
