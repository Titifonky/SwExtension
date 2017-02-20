using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ModuleMarchePositionner
{
    public abstract class PageMarchePositionner : BoutonPMPManager
    {
        protected CtrlCheckBox _CheckBox_EnregistrerSelection;

        protected void SvgNomComposant(Object SelBox, Parametre Param)
        {
            Component2 Cp = App.ModelDoc2.eSelect_RecupererComposant(1, ((CtrlSelectionBox)SelBox).Marque);
            if (Cp.IsRef() && _CheckBox_EnregistrerSelection.IsChecked)
                Param.SetValeur(Cp.eNomSansExt());
        }

        protected void SvgNomFonction(Object SelBox, Parametre Param)
        {
            Feature F = App.ModelDoc2.eSelect_RecupererObjet<Feature>(1, ((CtrlSelectionBox)SelBox).Marque);
            if (F.IsRef() && _CheckBox_EnregistrerSelection.IsChecked)
                Param.SetValeur(F.Name);
        }
    }
}
