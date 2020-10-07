﻿using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SwExtension;
using System;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Inclure/Exclure de la nomenclature"),
        ModuleNom("ExclureNomenclature")]
    public class BoutonExclureNomenclature : BoutonBase
    {
        private Parametre _pPropExclureNomenclature;

        public BoutonExclureNomenclature()
        {
            _pPropExclureNomenclature = _Config.AjouterParam("ExclureNomenclature", "ExcluNomenclature", "Nom de la propriete à verifier\n pour exclure de la nomenclature");
        }

        protected override void Command()
        {
            try
            {
                MdlBase.eRecParcourirComposants(InclureExclure, null);

                MdlBase.FeatureManager.UpdateFeatureTree();
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        private Boolean InclureExclure(Component2 cp)
        {
            if(!cp.IsSuppressed() && (eTypeDoc.Assemblage | eTypeDoc.Piece).HasFlag(cp.TypeDoc()))
            {
                if (cp.eModelDoc2().ePropExiste(_pPropExclureNomenclature.GetValeur<String>()))
                {
                    int MdlIntVal = cp.eModelDoc2().eGetProp(_pPropExclureNomenclature.GetValeur<String>()).eToInteger();

                    switch (MdlIntVal)
                    {
                        case 0:
                            {
                                cp.ExcludeFromBOM = false;
                                break;
                            }

                        case 1:
                            {
                                cp.ExcludeFromBOM = true;
                                //WindowLog.Ecrire(cp.Name2 + " : Est exclu");
                                break;
                            }

                        case 2:
                            {
                                if (cp.ePropExiste(_pPropExclureNomenclature.GetValeur<String>()))
                                {
                                    int CfgIntVal = cp.eProp(_pPropExclureNomenclature.GetValeur<String>()).eToInteger();

                                    switch (CfgIntVal)
                                    {
                                        case 0:
                                            {
                                                cp.ExcludeFromBOM = false;
                                                break;
                                            }

                                        case 1:
                                            {
                                                cp.ExcludeFromBOM = true;
                                                //WindowLog.Ecrire(cp.Name2);
                                                //WindowLog.Ecrire(cp.eNomConfiguration());
                                                //WindowLog.EcrireF("{0}  \"{1}\" : Est exclu", cp.Name2, cp.eNomConfiguration());
                                                break;
                                            }

                                        default:
                                            break;
                                    }

                                }
                                break;
                            }

                        default:
                            break;
                    }

                }


            }
            return false;
        }
    }
}
