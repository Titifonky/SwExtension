using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Linq;
using System.Reflection;

namespace SwExtension
{
    public abstract class BoutonPMPManager : BoutonBase, IPropertyManagerPage2Handler9
    {
        private Boolean _IsShow = false;
        public Boolean IsShow { get { return _IsShow; } }

        protected PropertyManagerPage2 _PmPage;
        protected Calque _Calque;

        public BoutonPMPManager()
        {
            try
            {
                SauverConfigBouton = false;

                var Options = PageOptions.Defaut;

                PageOptions pageOptionsAtt = GetType().GetCustomAttribute<PageOptions>();
                if (pageOptionsAtt.IsRef())
                    Options |= pageOptionsAtt.Val;

                int iErrors = 0;
                _PmPage = (PropertyManagerPage2)App.Sw.CreatePropertyManagerPage(TitreModule, (int)Options, this, ref iErrors);

                if (iErrors != (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
                {
                    this.LogMethode(new Object[] { "Erreur de création :", (swPropertyManagerPageStatus_e)iErrors });
                    IsInit = false;
                }
            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { "Erreur :", e });
                IsInit = false;
            }
        }

        protected override void Command()
        {
            if (!IsInit) return;

            try
            {
                _Calque = new Calque(_PmPage, NomModule);
                _Calque.Entete("Info", DescriptionModule);

                if (OnCalque.IsRef())
                    OnCalque();

                if (OnPreSelection.IsRef())
                    OnPreSelection();

                swPropertyManagerPageStatus_e r = (swPropertyManagerPageStatus_e)_PmPage.Show2(0);
                _IsShow = true;

                if (r == swPropertyManagerPageStatus_e.swPropertyManagerPage_CreationFailure)
                {
                    this.LogMethode(new Object[] { "Erreur de création" });
                    _IsShow = false;
                }
                else
                { }
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected delegate void OnCalqueEventHandler();
        protected event OnCalqueEventHandler OnCalque;

        protected delegate void OnPreSelectionEventHandler();
        protected event OnPreSelectionEventHandler OnPreSelection;

        protected delegate void OnRunAfterActivationEventHandler();
        protected event OnRunAfterActivationEventHandler OnRunAfterActivation;

        protected delegate void OnRunOnCloseEventHandler();
        protected event OnRunOnCloseEventHandler OnRunOnClose;

        protected delegate void OnRunAfterCloseEventHandler();
        protected event OnRunAfterCloseEventHandler OnRunAfterClose;

        protected delegate void OnRunOkCommandEventHandler();
        protected event OnRunOkCommandEventHandler OnRunOkCommand;

        protected delegate void OnRunCancelCommandEventHandler();
        protected event OnRunCancelCommandEventHandler OnRunCancelCommand;

        protected delegate void OnHelpEventHandler();
        protected event OnHelpEventHandler OnHelp;

        protected delegate Boolean OnNextPageEventHandler();
        protected event OnNextPageEventHandler OnNextPage;

        protected delegate Boolean OnPreviousPageEventHandler();
        protected event OnPreviousPageEventHandler OnPreviousPage;

        private Boolean CanRaiseEvent([CallerMemberName] String methode = "")
        {
            //this.LogMethode(new Object[] { "CanRaiseEvent", IsShow & IsInit });
            //Log.Message("Callback methode -> " + methode);
            return IsShow & IsInit;
        }

        #region Filtre et Test de selection

        /// <summary>
        /// Selectionne tous les composants parent jusqu'au composant de 1er niveau
        /// </summary>
        /// <param name="SelBox"></param>
        /// <param name="selection"></param>
        /// <param name="selType"></param>
        /// <param name="itemText"></param>
        /// <returns></returns>
        protected static Boolean SelectionnerComposantsParent(Object SelBox, Object selection, int selType, String itemText)
        {
            Component2 Cp = selection as Component2;

            if (Cp.IsRef())
            {
                List<Component2> Liste = Cp.eListeComposantParent();
                Liste.Insert(0, Cp);

                App.ModelDoc2.eSelectMulti(Liste, ((CtrlSelectionBox)SelBox).Marque, true);
            }

            return false;
        }

        /// <summary>
        /// Selectionne seulement une pièce
        /// </summary>
        /// <param name="SelBox"></param>
        /// <param name="selection"></param>
        /// <param name="selType"></param>
        /// <param name="itemText"></param>
        /// <returns></returns>
        protected static Boolean SelectionnerPiece(Object SelBox, Object selection, int selType, String itemText)
        {
            Component2 Cp = selection as Component2;
            if (Cp.TypeDoc() == eTypeDoc.Piece)
            {
                if (selType == (int)swSelectType_e.swSelCOMPONENTS)
                    return true;

                if (Cp.IsRef())
                    App.ModelDoc2.eSelectMulti(Cp, ((CtrlSelectionBox)SelBox).Marque, true);
            }

            return false;
        }

        /// <summary>
        /// Selectionne seulement un assemblage
        /// </summary>
        /// <param name="SelBox"></param>
        /// <param name="selection"></param>
        /// <param name="selType"></param>
        /// <param name="itemText"></param>
        /// <returns></returns>
        protected static Boolean SelectionnerAssemblage(Object SelBox, Object selection, int selType, String itemText)
        {
            Component2 Cp = selection as Component2;
            if (Cp.TypeDoc() == eTypeDoc.Assemblage)
            {
                if (selType == (int)swSelectType_e.swSelCOMPONENTS)
                    return true;

                if (Cp.IsRef())
                    App.ModelDoc2.eSelectMulti(Cp, ((CtrlSelectionBox)SelBox).Marque, true);
            }

            return false;
        }

        /// <summary>
        /// Selectionne le composant de 1er niveau
        /// </summary>
        /// <param name="SelBox"></param>
        /// <param name="selection"></param>
        /// <param name="selType"></param>
        /// <param name="itemText"></param>
        /// <returns></returns>
        protected static Boolean SelectionnerComposant1erNvx(Object SelBox, Object selection, int selType, String itemText)
        {
            Component2 Cp = selection as Component2;

            if (Cp.IsRef())
            {
                List<Component2> Liste = Cp.eListeComposantParent();

                if (Liste.Count == 0)
                    return true;

                CtrlSelectionBox box = SelBox as CtrlSelectionBox;

                Cp = Liste.Last();

                if (App.ModelDoc2.eSelect_RecupererListeComposants(box.Marque).Contains(Cp))
                    Cp.eDeSelectById(App.ModelDoc2);
                else
                    App.ModelDoc2.eSelectMulti(Cp, box.Marque, true);
            }

            return false;
        }

        /// <summary>
        /// Test si une chaine de caractère est un entier
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        protected static Boolean ValiderTextIsInteger(String text)
        {
            if (!text.Trim().eIsInteger())
                return false;

            return true;
        }

        /// <summary>
        /// Test si une chaine de caractère est un nombre decimal
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        protected static Boolean ValiderTextIsDouble(String text)
        {
            if (!text.Trim().eIsDouble())
                return false;

            return true;
        }

        /// <summary>
        /// Test si une chaine de caractère est un entier ou composé d'entiers séparés par des espaces
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        protected static Boolean ValiderListeTextIsInteger(String text)
        {
            List<String> Liste = new List<string>(text.Split(' '));
            foreach (String t in Liste)
            {
                if (!t.Trim().eIsInteger())
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Test si une chaine de caractère est un réel ou composé de réels séparés par des espaces
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        protected static Boolean ValiderListeTextIsDouble(String text)
        {
            List<String> Liste = new List<string>(text.Split(' '));
            foreach (String t in Liste)
            {
                if (!t.Trim().eIsDouble())
                    return false;
            }

            return true;
        }

        #endregion

        #region Evenements configurés

        private Boolean _Ok = false;

        void IPropertyManagerPage2Handler9.AfterActivation()
        {
            if (OnRunAfterActivation.IsRef())
                OnRunAfterActivation();
        }

        void IPropertyManagerPage2Handler9.OnClose(int Reason)
        {
            SauverConfig();

            if (OnRunOnClose.IsRef())
                OnRunOnClose();

            if (IsInit && (Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay))
                _Ok = true;
        }

        void IPropertyManagerPage2Handler9.AfterClose()
        {
            SauverConfigBouton = true;

            SauverConfig();

            if (OnRunAfterClose.IsRef())
                OnRunAfterClose();

            if (_Ok)
                if (OnRunOkCommand.IsRef())
                    OnRunOkCommand();
                else
                    if (OnRunCancelCommand.IsRef())
                    OnRunCancelCommand();

            //SauverConfig();

            _Config = null;
            _Calque = null;
            _PmPage = null;
        }

        bool IPropertyManagerPage2Handler9.OnHelp()
        {
            if (OnHelp.IsRef())
            {
                OnHelp();
            }

            return true;
        }

        bool IPropertyManagerPage2Handler9.OnNextPage()
        {
            if (OnNextPage.IsRef())
            {
                return OnNextPage();
            }

            return true;
        }

        bool IPropertyManagerPage2Handler9.OnPreviousPage()
        {
            if (OnPreviousPage.IsRef())
            {
                return OnPreviousPage();
            }

            return true;
        }

        void IPropertyManagerPage2Handler9.OnCheckboxCheck(int Id, bool Checked)
        {
            if (!CanRaiseEvent()) return;

            CtrlCheckBox CheckBox = _Calque.DicControl[Id] as CtrlCheckBox;
            if (CheckBox.IsRef())
            {
                CheckBox.IsChecked = Checked;
            }
        }

        void IPropertyManagerPage2Handler9.OnOptionCheck(int Id)
        {
            if (!CanRaiseEvent()) return;

            CtrlOption Option = _Calque.DicControl[Id] as CtrlOption;
            if (Option.IsRef())
            {
                Option.IsChecked = true;
            }
        }

        void IPropertyManagerPage2Handler9.OnGroupCheck(int Id, bool Checked)
        {
            if (!CanRaiseEvent()) return;

            GroupeAvecCheckBox GroupeCheckBox = _Calque.DicGroup[Id] as GroupeAvecCheckBox;
            if (GroupeCheckBox.IsRef())
            {
                GroupeCheckBox.IsChecked = Checked;
            }
        }

        void IPropertyManagerPage2Handler9.OnGroupExpand(int Id, bool Expanded)
        {
            if (!CanRaiseEvent()) return;

            Groupe Groupe = _Calque.DicGroup[Id] as Groupe;
            if (Groupe.IsRef())
            {
                Groupe.Expanded = Expanded;
            }
        }

        void IPropertyManagerPage2Handler9.OnSelectionboxListChanged(int Id, int Count)
        {
            if (!CanRaiseEvent()) return;

            CtrlSelectionBox Box = _Calque.DicControl[Id] as CtrlSelectionBox;
            Box.SelectionChanged(Box, Count);
        }

        void IPropertyManagerPage2Handler9.OnSelectionboxFocusChanged(int Id)
        {
            if (!CanRaiseEvent()) return;

            CtrlSelectionBox Box = _Calque.DicControl[Id] as CtrlSelectionBox;
            if (Box.IsRef())
            {
                Box.SelectionboxFocusChanged(Box);
            }
        }

        bool IPropertyManagerPage2Handler9.OnSubmitSelection(int Id, object Selection, int SelType, ref string ItemText)
        {
            if (!CanRaiseEvent()) return false;

            CtrlSelectionBox Box = _Calque.DicControl[Id] as CtrlSelectionBox;
            if (Box.IsRef())
            {
                return Box.SubmitSelection(Box, Selection, SelType, ItemText);
            }

            return false;
        }

        void IPropertyManagerPage2Handler9.OnTextboxChanged(int Id, string Text)
        {
            if (!CanRaiseEvent()) return;

            CtrlTextBox TxtBox = _Calque.DicControl[Id] as CtrlTextBox;
            if (TxtBox.IsRef())
            {
                TxtBox.TextBoxChanged(Text);
            }
        }

        void IPropertyManagerPage2Handler9.OnComboboxEditChanged(int Id, string Text)
        {
            if (!CanRaiseEvent()) return;

            CtrlTextComboBox CmbBox = _Calque.DicControl[Id] as CtrlTextComboBox;
            if (CmbBox.IsRef())
            {
                CmbBox.ComboboxEditChanged(CmbBox, Text);
            }
        }

        void IPropertyManagerPage2Handler9.OnComboboxSelectionChanged(int Id, int Item)
        {
            if (!CanRaiseEvent()) return;

            CtrlBaseComboBox CmbBox = _Calque.DicControl[Id] as CtrlBaseComboBox;
            if (CmbBox.IsRef())
            {
                CmbBox.SelectionChanged(CmbBox, Item);
            }
        }

        void IPropertyManagerPage2Handler9.OnListboxSelectionChanged(int Id, int Item)
        {
            if (!CanRaiseEvent()) return;

            CtrlBaseListBox ListBox = _Calque.DicControl[Id] as CtrlBaseListBox;
            if (ListBox.IsRef())
            {
                if (ListBox.swListBox.GetSelectedItemsCount() > 0)
                    Item = ListBox.swListBox.GetSelectedItems()[0];
                
                ListBox.SelectionChanged(ListBox, Item);
            }
        }

        void IPropertyManagerPage2Handler9.OnButtonPress(int Id)
        {
            if (!CanRaiseEvent()) return;

            CtrlButton Bouton = _Calque.DicControl[Id] as CtrlButton;
            if (Bouton.IsRef())
            {
                Bouton.ButtonPress(Bouton);
            }
        }

        void IPropertyManagerPage2Handler9.OnGainedFocus(int Id)
        {
            if (!CanRaiseEvent()) return;

            Control Control = _Calque.DicControl[Id];
            if (Control.IsRef())
            {
                Control.GainedFocus(Control);
            }
        }

        void IPropertyManagerPage2Handler9.OnLostFocus(int Id)
        {
            if (!CanRaiseEvent()) return;

            Control Control = _Calque.DicControl[Id];
            if (Control.IsRef())
            {
                Control.LostFocus(Control);
            }
        }

        #endregion

        int IPropertyManagerPage2Handler9.OnActiveXControlCreated(int Id, bool Status) { return 0; }

        bool IPropertyManagerPage2Handler9.OnKeystroke(int Wparam, int Message, int Lparam, int Id) { return true; }

        void IPropertyManagerPage2Handler9.OnListboxRMBUp(int Id, int PosX, int PosY) { }

        void IPropertyManagerPage2Handler9.OnNumberBoxTrackingCompleted(int Id, double Value) { }

        void IPropertyManagerPage2Handler9.OnNumberboxChanged(int Id, double Value) { }

        void IPropertyManagerPage2Handler9.OnPopupMenuItem(int Id) { }

        void IPropertyManagerPage2Handler9.OnPopupMenuItemUpdate(int Id, ref int retval) { }

        bool IPropertyManagerPage2Handler9.OnTabClicked(int Id) { return true; }

        bool IPropertyManagerPage2Handler9.OnPreview() { return true; }

        void IPropertyManagerPage2Handler9.OnRedo() { }

        void IPropertyManagerPage2Handler9.OnUndo() { }

        void IPropertyManagerPage2Handler9.OnSelectionboxCalloutCreated(int Id) { }

        void IPropertyManagerPage2Handler9.OnSelectionboxCalloutDestroyed(int Id) { }

        void IPropertyManagerPage2Handler9.OnSliderPositionChanged(int Id, double Value) { }

        void IPropertyManagerPage2Handler9.OnSliderTrackingCompleted(int Id, double Value) { }

        void IPropertyManagerPage2Handler9.OnWhatsNew() { }

        int IPropertyManagerPage2Handler9.OnWindowFromHandleControlCreated(int Id, bool Status) { return 0; }
    }

    public class PageOptions : System.Attribute
    {
        public static readonly swPropertyManagerPageOptions_e Defaut = swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton |
                                    swPropertyManagerPageOptions_e.swPropertyManagerOptions_CancelButton |
                                    swPropertyManagerPageOptions_e.swPropertyManagerOptions_LockedPage
                                    ;

        private swPropertyManagerPageOptions_e _Val;

        public PageOptions(swPropertyManagerPageOptions_e val) { _Val = val; }

        public swPropertyManagerPageOptions_e Val { get { return _Val; } }
    }
}
