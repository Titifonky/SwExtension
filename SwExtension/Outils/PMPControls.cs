using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;

namespace Outils
{
    public static class ID
    {
        public static void Reset()
        {
            _idGroup = 1;
            _idControl = 1000;
            _marque = 0;
            _couleurCourante = 0;
        }

        private static int _idGroup = 1;

        public static int NextIdGroup { get { return _idGroup++; } }

        private static int _idControl = 1000;

        public static int NextIdControl { get { return _idControl++; } }

        private static int _marque = 0;

        public static int NextMarque { get { return (int)Math.Pow(2, _marque++); } }

        private static swUserPreferenceIntegerValue_e[] _couleur = new swUserPreferenceIntegerValue_e[] { swUserPreferenceIntegerValue_e.swSystemColorsSelectedItem1,
                                                                                                    swUserPreferenceIntegerValue_e.swSystemColorsSelectedItem2,
                                                                                                    swUserPreferenceIntegerValue_e.swSystemColorsSelectedItem3,
                                                                                                    swUserPreferenceIntegerValue_e.swSystemColorsSelectedItem4
                                                                                                      };
        private static int _couleurCourante = 0;
        public static swUserPreferenceIntegerValue_e NextCouleur
        {
            get
            {
                if (_couleurCourante == _couleur.Length)
                    _couleurCourante = 0;

                return _couleur[_couleurCourante++];
            }
        }
    }

    public class Calque
    {
        private PropertyManagerPage2 _SwPage;

        private String _NomModule = "";

        public Dictionary<int, Groupe> DicGroup = new Dictionary<int, Groupe>();

        public Dictionary<int, Control> DicControl = new Dictionary<int, Control>();

        public PropertyManagerPage2 swPage { get { return _SwPage; } }

        public String NomModuleSvgParametre { get { return _NomModule; } }

        public Calque(PropertyManagerPage2 swPage, String nomModuleSvgParametre = "")
        {
            _SwPage = swPage;
            _NomModule = nomModuleSvgParametre;
            ID.Reset();
        }

        public void Entete(String titre, String message)
        {
            _SwPage.SetMessage3(message,
                                (int)swPropertyManagerPageMessageVisibility.swMessageBoxVisible,
                                (int)swPropertyManagerPageMessageExpanded.swMessageBoxExpand,
                                titre);
        }

        public void CacherEntete()
        {
            _SwPage.SetMessage3("",
                                (int)swPropertyManagerPageMessageVisibility.swMessageBoxHidden,
                                (int)swPropertyManagerPageMessageExpanded.swMessageBoxCompress,
                                "");
        }

        private readonly int Option = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible;

        public Groupe AjouterGroupe(String titre)
        {
            return new Groupe(this, Option + (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded, titre);
        }

        public Groupe AjouterGroupe(Parametre param)
        {
            return new Groupe(this, Option + (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded, param);
        }

        public GroupeAvecCheckBox AjouterGroupeAvecCheckBox(String titre)
        {
            return new GroupeAvecCheckBox(this, Option + (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Checkbox, titre);
        }

        public GroupeAvecCheckBox AjouterGroupeAvecCheckBox(Parametre param)
        {
            int Option = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible +
                            (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Checkbox;

            return new GroupeAvecCheckBox(this, Option, param); ;
        }

        ~Calque()
        {
            DicGroup.Clear();
            DicGroup = null;
            DicControl.Clear();
            DicControl = null;

            _SwPage = null;
        }

    }

    public abstract class Base
    {
        public Parametre Param { get; set; }

        protected void SetParametre<T>(T objet)
        {
            if (Param == null) return;

            Param.SetValeur(objet);
        }

        protected T GetParametre<T>()
        {
            if (Param == null) return default(T);

            return Param.GetValeur<T>();
        }
    }

    public class Groupe : Base
    {
        protected int _Id;

        protected Calque _Page;

        public Calque Page { get { return _Page; } }

        protected PropertyManagerPageGroup _swGroup;

        public PropertyManagerPageGroup swGroup { get { return _swGroup; } }

        public int Id { get { return _Id; } }

        public String Titre
        {
            get
            {
                return _swGroup.Caption;
            }
            set
            {
                _swGroup.Caption = value;
            }
        }

        private Boolean _Expanded = false;

        public Boolean Expanded
        {
            get { return _Expanded; }
            set { if (value) Expand(); else UnExpand(); }
        }

        private Boolean _Visible = true;

        public Boolean Visible
        {
            get { return _Visible; }
            set { _swGroup.Visible = value; _Visible = value; }
        }

        public Groupe(Calque page, int options, String titre)
        {
            _Page = page;
            _Id = ID.NextIdGroup;
            _swGroup = _Page.swPage.AddGroupBox(_Id, titre, options);
            _Page.DicGroup.Add(_Id, this);
        }

        public Groupe(Calque page, int options, Parametre param)
        {
            _Page = page;
            _Id = ID.NextIdGroup;
            _swGroup = _Page.swPage.AddGroupBox(_Id, param.Intitule, options);
            _Page.DicGroup.Add(_Id, this);
            Param = param;

            Expanded = GetParametre<Boolean>();
        }

        public delegate void OnExpandEventHandler();

        public event OnExpandEventHandler OnExpand;

        public void Expand()
        {
            _swGroup.Expanded = true;
            _Expanded = true;

            if (OnExpand.IsRef())
                OnExpand();
        }

        public delegate void OnUnExpandEventHandler();

        public event OnUnExpandEventHandler OnUnExpand;

        public void UnExpand()
        {
            _swGroup.Expanded = false;
            _Expanded = false;

            if (OnUnExpand.IsRef())
                OnUnExpand();
        }

        private readonly int Option = (int)swAddControlOptions_e.swControlOptions_Visible + (int)swAddControlOptions_e.swControlOptions_Enabled;

        public CtrlSelectionBox AjouterSelectionBox(String labelTitre, String labelTip = "", Boolean avecCouleur = true)
        {
            return new CtrlSelectionBox(this, Option, avecCouleur, labelTitre, labelTip);
        }

        public CtrlCheckBox AjouterCheckBox(String titre)
        {
            return new CtrlCheckBox(this, titre, Option);
        }

        public CtrlCheckBox AjouterCheckBox(Parametre param)
        {
            return new CtrlCheckBox(this, Option, param);
        }

        public CtrlOption AjouterOption(String titre)
        {
            var obj = new CtrlOption(this, titre, Option, _ListeCtrlOption);
            return obj;
        }

        public CtrlOption AjouterOption(Parametre param)
        {
            var obj = new CtrlOption(this, Option, param, _ListeCtrlOption);
            return obj;
        }

        public CtrlButton AjouterBouton(String titre, String tip = "")
        {
            return new CtrlButton(this, titre, Option, tip);
        }

        public CtrlLabel AjouterLabel(String titre, String tip = "")
        {
            return new CtrlLabel(this, titre, Option, tip);
        }

        public CtrlImage AjouterImage(String titre, String tip = "")
        {
            return new CtrlImage(this, titre, Option, tip);
        }

        public CtrlTextBox AjouterTexteBox(String labelTitre = "", String labelTip = "")
        {
            return new CtrlTextBox(this, Option, labelTitre, labelTip);
        }

        public CtrlTextBox AjouterTexteBox(Parametre param, Boolean AvecIntitule = false)
        {
            return new CtrlTextBox(this, Option, AvecIntitule, param);
        }

        public CtrlTextComboBox AjouterTextComboBox(String labelTitre = "", String labelTip = "")
        {
            return new CtrlTextComboBox(this, Option, labelTitre, labelTip);
        }

        public CtrlTextComboBox AjouterTextComboBox(Parametre param, Boolean AvecIntitule = true)
        {
            return new CtrlTextComboBox(this, Option, AvecIntitule, param);
        }

        public CtrlEnumComboBox<T, D> AjouterEnumComboBox<T, D>(Parametre param, Boolean AvecIntitule = true)
            where T : struct, IComparable, IFormattable, IConvertible
            where D : PersoAttribute
        {
            return new CtrlEnumComboBox<T, D>(this, Option, AvecIntitule, param);
        }

        public CtrlTextListBox AjouterTextListBox(String labelTitre = "", String labelTip = "")
        {
            return new CtrlTextListBox(this, Option, labelTitre, labelTip);
        }

        public CtrlTextListBox AjouterTextListBox(Parametre param, Boolean AvecIntitule = true)
        {
            return new CtrlTextListBox(this, Option, AvecIntitule, param);
        }

        private List<CtrlOption> _ListeCtrlOption = new List<CtrlOption>();
    }

    public class GroupeAvecCheckBox : Groupe
    {
        private Boolean _IsChecked = false;

        public Boolean IsChecked
        {
            get { return _IsChecked; }
            set { IsCheck(null, value); }
        }

        public GroupeAvecCheckBox(Calque page, int options, String titre)
            : base(page, options, titre)
        { }

        public GroupeAvecCheckBox(Calque page, int options, Parametre param)
            : base(page, options, param)
        {
            IsChecked = GetParametre<Boolean>();
        }

        public delegate void OnIsCheckEventHandler(Object sender, Boolean value);

        public event OnIsCheckEventHandler OnIsCheck;

        public void IsCheck(Object sender, Boolean value)
        {
            if (OnIsCheck.IsRef())
                OnIsCheck(this, value);

            if (value)
                Check(this);
            else
                UnCheck(this);
        }

        public delegate void OnCheckEventHandler(Object sender);

        public event OnCheckEventHandler OnCheck;

        public void Check(Object sender)
        {
            _swGroup.Checked = true;
            _IsChecked = true;
            SetParametre(true);

            if (OnCheck.IsRef())
                OnCheck(this);
        }

        public delegate void OnUnCheckEventHandler(Object sender);

        public event OnUnCheckEventHandler OnUnCheck;

        public void UnCheck(Object sender)
        {
            _swGroup.Checked = false;
            _IsChecked = false;
            SetParametre(false);

            if (OnUnCheck.IsRef())
                OnUnCheck(this);
        }
    }

    public abstract class Control : Base
    {
        private int _Id;

        protected Boolean _Init = false;

        protected Groupe _Groupe;

        protected PropertyManagerPageControl _swControl;

        public PropertyManagerPageControl swControl { get { return _swControl; } }

        protected PropertyManagerPageLabel _swLabelBase;

        public PropertyManagerPageLabel swLabelBase { get { return _swLabelBase; } }

        public Boolean IsEnabled { get { return _swControl.Enabled; } set { IsEnable(this, value); } }

        public Boolean Visible { get { return _swControl.Visible; } set { IsVisible(this, value); } }

        public String LabelText
        {
            get
            {
                if (_swLabelBase.IsRef())
                    return _swLabelBase.Caption;

                return null;
            }

            set
            {
                if (_swLabelBase.IsRef())
                    _swLabelBase.Caption = value;
            }
        }

        public virtual String Caption
        {
            get { return ""; }
            set { }
        }

        public int Id { get { return _Id; } protected set { _Id = value; } }

        public short Indent
        {
            set
            {
                if (_swLabelBase.IsRef())
                {
                    PropertyManagerPageControl c = _swLabelBase as PropertyManagerPageControl;
                    c.Left = value;
                }

                _swControl.Left = value;
            }
        }

        public void StdIndent(Boolean ControlOnly = false)
        {
            if (ControlOnly)
                _swControl.Left = 10;
            else
                Indent = 10;
        }

        public short Top
        {
            set
            {
                if (_swLabelBase.IsRef())
                {
                    PropertyManagerPageControl c = _swLabelBase as PropertyManagerPageControl;
                    c.Top = value;
                }

                _swControl.Top = value;
            }
        }

        public void BackgroundColor(Color couleur, Boolean ControlOnly = false)
        {
            int cl16 = couleur.R | couleur.G << 5 | couleur.B << 11;

            if (!ControlOnly && _swLabelBase.IsRef())
            {
                PropertyManagerPageControl c = _swLabelBase as PropertyManagerPageControl;
                c.BackgroundColor = cl16;
            }

            _swControl.BackgroundColor = cl16;
        }

        private void Init(Groupe groupe, Boolean avecLabel, String labelTitre, String labelTip)
        {
            _Init = true;
            _Groupe = groupe;
            _Id = ID.NextIdControl;

            if (avecLabel)
            {
                _swLabelBase = _Groupe.swGroup.AddControl2(ID.NextIdControl, (int)swPropertyManagerPageControlType_e.swControlType_Label,
                                                                                        labelTitre,
                                                                                        (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge,
                                                                                        (int)swAddControlOptions_e.swControlOptions_Visible + (int)swAddControlOptions_e.swControlOptions_Enabled,
                                                                                        labelTip);
            }
        }

        public Control(Groupe groupe, String labelTitre = "", String labelTip = "")
        {
            Init(groupe, !String.IsNullOrWhiteSpace(labelTitre), labelTitre, labelTip);
        }

        public Control(Groupe groupe, Parametre param, Boolean avecLabel)
        {
            Param = param;
            Init(groupe, avecLabel, param.Intitule, param.Tip);
        }

        public virtual void ApplyParam() { }

        public Boolean Focus { get { return _Groupe.Page.swPage.GetFocus() == _Id; } set { if (value) GainedFocus(this); } }

        public delegate void OnIsVisibleHandler(Object sender, Boolean value);

        public event OnIsVisibleHandler OnIsVisible;

        public void IsVisible(Object sender, Boolean value)
        {
            if (OnIsVisible.IsRef())
                OnIsVisible(this, value);

            _swControl.Visible = value;
            if (_swLabelBase.IsRef())
                ((PropertyManagerPageControl)_swLabelBase).Visible = value;
        }

        public delegate void OnIsEnableHandler(Object sender, Boolean value);

        public event OnIsEnableHandler OnIsEnable;

        public void IsEnable(Object sender, Boolean value)
        {
            if (OnIsEnable.IsRef())
                OnIsEnable(this, value);

            if (OnIsDisable.IsRef())
                OnIsDisable(this, !value);

            if (value)
                Enable(sender);
            else
                Disable(sender);
        }

        public delegate void OnIsDisableHandler(Object sender, Boolean value);

        public event OnIsEnableHandler OnIsDisable;

        public void IsDisable(Object sender, Boolean value)
        {
            if (OnIsDisable.IsRef())
                OnIsDisable(this, value);

            if (OnIsEnable.IsRef())
                OnIsEnable(this, !value);

            if (value)
                Disable(sender);
            else
                Enable(sender);
        }

        public delegate void OnEnableHandler(Object sender);

        public event OnEnableHandler OnEnable;

        public void Enable(Object sender)
        {
            _swControl.Enabled = true;
            if (_swLabelBase.IsRef())
            {
                PropertyManagerPageControl ctrl = _swLabelBase as PropertyManagerPageControl;
                ctrl.Enabled = true;
            }

            if (OnEnable.IsRef())
                OnEnable(this);
        }

        public delegate void OnDisableHandler(Object sender);

        public event OnDisableHandler OnDisable;

        public void Disable(Object sender)
        {
            _swControl.Enabled = false;
            if (_swLabelBase.IsRef())
            {
                PropertyManagerPageControl ctrl = _swLabelBase as PropertyManagerPageControl;
                ctrl.Enabled = false;
            }

            if (OnDisable.IsRef())
                OnDisable(this);
        }

        public delegate void OnGainedFocusHandler(Object sender);

        public event OnGainedFocusHandler OnGainedFocus;

        public void GainedFocus(Object sender)
        {
            if (!this.Focus)
                _Groupe.Page.swPage.SetFocus(Id);

            if (OnGainedFocus.IsRef())
                OnGainedFocus(this);
        }

        public delegate void OnLostFocusEventHandler(Object sender);

        public event OnLostFocusEventHandler OnLostFocus;

        public void LostFocus(Object sender)
        {
            if (OnLostFocus.IsRef())
                OnLostFocus(this);
        }
    }

    public class CtrlImage : Control
    {
        private PropertyManagerPageBitmap _swImage;

        public PropertyManagerPageBitmap swImage { get { return _swImage; } }

        public Boolean Chemin(String cheminCouleur, String cheminMasque)
        {
            try
            {
                return swImage.SetBitmapByName(cheminCouleur, cheminMasque);
            }
            catch (Exception e) { LogDebugging.Log.Message( e ); }

            return false;
        }

        private void Init(String intitule, int options, String tip = "")
        {
            _swImage = _Groupe.swGroup.AddControl2(Id, (int)swPropertyManagerPageControlType_e.swControlType_Bitmap,
                                                                                        intitule,
                                                                                        (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent,
                                                                                        options,
                                                                                        tip);
            _swControl = (PropertyManagerPageControl)_swImage;

            _Groupe.Page.DicControl.Add(Id, this);
        }

        public CtrlImage(Groupe groupe, String info, int options, String tip)
            : base(groupe)
        {
            Init(info, options, tip);
        }
    }

    public class CtrlLabel : Control
    {
        private PropertyManagerPageLabel _swLabel;

        public PropertyManagerPageLabel swLabel { get { return _swLabel; } }

        public override String Caption { get { return _swLabel.Caption; } set { _swLabel.Caption = value; } }

        private void Init(String intitule, int options, String tip = "")
        {
            _swLabel = _Groupe.swGroup.AddControl2(Id, (int)swPropertyManagerPageControlType_e.swControlType_Label,
                                                                                        intitule,
                                                                                        (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent,
                                                                                        options,
                                                                                        tip);
            _swControl = (PropertyManagerPageControl)_swLabel;

            _Groupe.Page.DicControl.Add(Id, this);
        }

        public CtrlLabel(Groupe groupe, String info, int options, String tip)
            : base(groupe)
        {
            Init(info, options, tip);
        }
    }

    public class CtrlButton : Control
    {
        private PropertyManagerPageButton _swButton;

        public PropertyManagerPageButton swButton { get { return _swButton; } }

        public override String Caption { get { return _swButton.Caption; } set { _swButton.Caption = value; } }

        private void Init(String intitule, int options, String tip = "")
        {
            _swButton = _Groupe.swGroup.AddControl2(Id, (int)swPropertyManagerPageControlType_e.swControlType_Button,
                                                                                        intitule,
                                                                                        (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent,
                                                                                        options,
                                                                                        tip);
            _swControl = (PropertyManagerPageControl)_swButton;

            _Groupe.Page.DicControl.Add(Id, this);
        }

        public CtrlButton(Groupe groupe, String info, int options, String tip)
            : base(groupe)
        {
            Init(info, options, tip);
        }

        public delegate void OnButtonPressEventHandler(Object sender);

        public event OnButtonPressEventHandler OnButtonPress;

        public void ButtonPress(Object sender)
        {
            if (OnButtonPress.IsRef())
                OnButtonPress(this);
        }
    }

    public class CtrlSelectionBox : Control
    {
        private PropertyManagerPageSelectionbox _swSelectionBox;

        public PropertyManagerPageSelectionbox swSelectionBox { get { return _swSelectionBox; } }

        public Boolean UneSeuleEntite { get { return _swSelectionBox.SingleEntityOnly; } set { _swSelectionBox.SingleEntityOnly = value; } }

        public Boolean SelectionMultipleMemeEntite { get { return _swSelectionBox.AllowMultipleSelectOfSameEntity; } set { _swSelectionBox.AllowMultipleSelectOfSameEntity = value; } }

        public Boolean SelectionDansMultipleBox { get { return _swSelectionBox.AllowSelectInMultipleBoxes; } set { _swSelectionBox.AllowSelectInMultipleBoxes = value; } }

        public short Height { get { return _swSelectionBox.Height; } set { _swSelectionBox.Height = value; } }

        public int Hauteur { set { _swSelectionBox.Height = (short)(value * 10); } }

        public int Nb { get { return _swSelectionBox.ItemCount; } }

        private int _Marque = 1;

        public int Marque { get { return _Marque; } set { _swSelectionBox.Mark = value; _Marque = value; } }

        private List<swSelectType_e> Liste;

        public void FiltreSelection(List<swSelectType_e> liste)
        {
            Liste = liste;
            int[] Filtre = new int[Liste.Count];
            for (int i = 0; i < Liste.Count; i++)
            {
                Filtre[i] = (int)Liste[i];
            }

            _swSelectionBox.SetSelectionFilters(Filtre);
        }

        public void FiltreSelection(params swSelectType_e[] liste)
        {
            FiltreSelection(liste.ToList());
        }

        private swUserPreferenceIntegerValue_e _Couleur;

        public swUserPreferenceIntegerValue_e Couleur
        {
            set
            {
                _Couleur = value;
                _swSelectionBox.SetSelectionColor(true, (int)value);
            }

            get
            {
                return _Couleur;
            }
        }

        private void Init(int options, Boolean avecCouleurs, String tip = "")
        {
            _swSelectionBox = _Groupe.swGroup.AddControl2(Id, (int)swPropertyManagerPageControlType_e.swControlType_Selectionbox,
                                                                                        tip,
                                                                                        (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge,
                                                                                        options,
                                                                                        tip);
            _swControl = (PropertyManagerPageControl)_swSelectionBox;
            Marque = ID.NextMarque;

            if (avecCouleurs)
                Couleur = ID.NextCouleur;

            _Groupe.Page.DicControl.Add(Id, this);
        }

        public CtrlSelectionBox(Groupe groupe, int options, Boolean avecCouleurs = true, String labelTitre = "", String labelTip = "")
            : base(groupe, labelTitre, labelTip)
        {
            Init(options, avecCouleurs, labelTip);
        }

        public delegate void ApplyOnSelectionEventHandler(Object sender);

        public event ApplyOnSelectionEventHandler OnApplyOnSelection;

        public void ApplyOnSelection(Object sender)
        {
            if (OnApplyOnSelection.IsRef())
                OnApplyOnSelection(this);
        }

        public delegate void OnSelectionChangedEventHandler(Object sender, int nb);

        public event OnSelectionChangedEventHandler OnSelectionChanged;

        public void SelectionChanged(Object sender, int nb)
        {
            if (OnSelectionChanged.IsRef())
                OnSelectionChanged(this, nb);

            if(nb > 0)
                ApplyOnSelection(sender);
        }

        public delegate void OnSelectionboxFocusChangedEventHandler(Object sender);

        public event OnSelectionboxFocusChangedEventHandler OnSelectionboxFocusChanged;

        public void SelectionboxFocusChanged(Object sender)
        {
            if (OnSelectionboxFocusChanged.IsRef())
                OnSelectionboxFocusChanged(this);
        }

        public delegate Boolean OnSubmitSelectionEventHandler(Object sender, Object selection, int SelType, String itemText);

        public OnSubmitSelectionEventHandler OnSubmitSelection;

        public Boolean SubmitSelection(Object sender, Object selection, int selType, String itemText)
        {
            if (OnSubmitSelection.IsRef())
                return OnSubmitSelection(this, selection, selType, itemText);

            return true;
        }
    }

    public class CtrlCheckBox : Control
    {
        private PropertyManagerPageCheckbox _swCheckBox;

        public PropertyManagerPageCheckbox swCheckBox { get { return _swCheckBox; } }

        private Boolean _IsChecked = false;

        public Boolean IsChecked
        {
            get { return _IsChecked; }
            set { IsCheck(this, value); }
        }

        public override String Caption { get { return _swCheckBox.Caption; } set { _swCheckBox.Caption = value; } }

        private void Init(String titre, int options, String tip = "")
        {
            _swCheckBox = _Groupe.swGroup.AddControl2(Id, (int)swPropertyManagerPageControlType_e.swControlType_Checkbox,
                                                                                        titre,
                                                                                        (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge,
                                                                                        options,
                                                                                        tip);
            _swControl = (PropertyManagerPageControl)_swCheckBox;

            _Groupe.Page.DicControl.Add(Id, this);
        }

        public CtrlCheckBox(Groupe groupe, String titre, int options, String labelTitre = "", String labelTip = "")
            : base(groupe, labelTitre, labelTip)
        {
            Init(titre, options, labelTip);

            _Init = false;
        }

        public CtrlCheckBox(Groupe groupe, int options, Parametre param, Boolean avecLabel = false)
            : base(groupe, param, avecLabel)
        {
            Init(param.Intitule, options, param.Tip);

            ApplyParam();

            _Init = false;
        }

        public override void ApplyParam()
        {
            _Init = true;
            if (Param.IsRef())
                IsCheck(this, GetParametre<Boolean>());
            _Init = false;
        }

        public delegate void OnIsCheckEventHandler(Object sender, Boolean value);

        public event OnIsCheckEventHandler OnIsCheck;

        public void IsCheck(Object sender, Boolean value)
        {
            ApplyValue(value);

            if (OnIsCheck.IsRef())
                OnIsCheck(this, value);

            if (value)
                Check(sender);
            else
                UnCheck(sender);
        }

        public delegate void OnCheckEventHandler(Object sender);

        public event OnCheckEventHandler OnCheck;

        public void Check(Object sender)
        {
            ApplyValue(true);

            if (OnCheck.IsRef())
                OnCheck(this);
        }

        public delegate void OnUnCheckEventHandler(Object sender);

        public event OnUnCheckEventHandler OnUnCheck;

        public void UnCheck(Object sender)
        {
            ApplyValue(false);

            if (OnUnCheck.IsRef())
                OnUnCheck(this);
        }

        private void ApplyValue(Boolean value)
        {
            _swCheckBox.Checked = value;

            _IsChecked = value;
            SetParametre(value);
        }
    }

    public class CtrlOption : Control
    {
        private List<CtrlOption> _ListeCtrlOption;

        private PropertyManagerPageOption _swOption;

        public PropertyManagerPageOption swOption { get { return _swOption; } }

        private Boolean _IsChecked = false;

        public Boolean IsChecked
        {
            get { return _IsChecked; }
            set
            {
                _Init = true;
                if (value)
                    Check(this);
                else
                    UnCheck();
                _Init = false;
            }
        }

        private void Init(String titre, int options, List<CtrlOption> ListeCtrlOption, String tip = "")
        {
            _ListeCtrlOption = ListeCtrlOption;
            _ListeCtrlOption.Add(this);

            _swOption = _Groupe.swGroup.AddControl2(Id, (int)swPropertyManagerPageControlType_e.swControlType_Option,
                                                                                        titre,
                                                                                        (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge,
                                                                                        options,
                                                                                        tip);
            _swControl = (PropertyManagerPageControl)_swOption;

            _Groupe.Page.DicControl.Add(Id, this);
        }

        public CtrlOption(Groupe groupe, String titre, int options, List<CtrlOption> ListeCtrlOption, String labelTitre = "", String labelTip = "")
            : base(groupe, labelTitre, labelTip)
        {

            Init(titre, options, ListeCtrlOption, labelTip);

            _Init = false;
        }

        public CtrlOption(Groupe groupe, int options, Parametre param, List<CtrlOption> ListeCtrlOption, Boolean avecLabel = false)
            : base(groupe, param, avecLabel)
        {
            Init(param.Intitule, options, ListeCtrlOption, param.Tip);

            ApplyParam();

            _Init = false;
        }

        public override void ApplyParam()
        {
            _Init = true;
            if (Param.IsRef())
            {
                var r = GetParametre<Boolean>();
                if(r)
                    Check(this);
            }
            _Init = false;
        }

        public delegate void OnCheckEventHandler(Object sender);

        public event OnCheckEventHandler OnCheck;

        public void Check(Object sender)
        {
            if(_Init)
                _swOption.Checked = true;

            _IsChecked = true;
            SetParametre(true);

            if (OnCheck.IsRef())
                OnCheck(this);

            foreach (var option in _ListeCtrlOption)
            {
                if (this.Id != option.Id)
                    option.UnCheck();
            }
        }

        public delegate void OnUnCheckEventHandler(Object sender);

        public event OnCheckEventHandler OnUnCheck;

        public void UnCheck()
        {
            SetParametre(false);
            _IsChecked = false;

            if (OnUnCheck.IsRef())
                OnUnCheck(this);
        }
    }

    public class CtrlTextBox : Control
    {
        private Boolean _EcraserTexte = false;

        private PropertyManagerPageTextbox _swTexteBox;

        public PropertyManagerPageTextbox swTexteBox { get { return _swTexteBox; } }

        private String _Text = "";

        public String Text {
            get { return _Text; }
            set {
                _EcraserTexte = true;
                TextBoxChanged(this, value);
                _EcraserTexte = false;
            }
        }

        public T GetTextAs<T>()
        {
            T Val;

            try
            {
                Val = (T)Convert.ChangeType(Text, typeof(T));
            }
            catch
            {
                Val = (T)typeof(T).GetDefaultValue();
            }

            return Val;
        }

        public void TextBoxChanged(String t)
        {
            TextBoxChanged(this, t);
        }

        public short Height { get { return _swTexteBox.Height; } set { _swTexteBox.Height = value; } }

        private Boolean GetStyle(swPropMgrPageTextBoxStyle_e p)
        {
            return (_swTexteBox.Style & (int)p) > 0;
        }

        private void SetStyle(Boolean etat, swPropMgrPageTextBoxStyle_e p)
        {
            int s = _swTexteBox.Style;
            if (etat)
                _swTexteBox.Style = s | (int)p;
            else if (GetStyle(p))
                _swTexteBox.Style = ((s - (int)p) > 0) ? (s - (int)p) : s;
        }

        public Boolean Multiligne
        {
            get
            {
                return GetStyle(swPropMgrPageTextBoxStyle_e.swPropMgrPageTextBoxStyle_Multiline);
            }
            set
            {
                SetStyle(value, swPropMgrPageTextBoxStyle_e.swPropMgrPageTextBoxStyle_Multiline);
            }
        }

        public Boolean SansBords
        {
            get
            {
                return GetStyle(swPropMgrPageTextBoxStyle_e.swPropMgrPageTextBoxStyle_NoBorder);
            }
            set
            {
                SetStyle(value, swPropMgrPageTextBoxStyle_e.swPropMgrPageTextBoxStyle_NoBorder);
            }
        }

        public Boolean LectureSeule
        {
            get
            {
                return GetStyle(swPropMgrPageTextBoxStyle_e.swPropMgrPageTextBoxStyle_ReadOnly);
            }
            set
            {
                SetStyle(value, swPropMgrPageTextBoxStyle_e.swPropMgrPageTextBoxStyle_ReadOnly);
            }
        }

        public Boolean NotifieSurFocus
        {
            get
            {
                return GetStyle(swPropMgrPageTextBoxStyle_e.swPropMgrPageTextBoxStyle_NotifyOnlyWhenFocusLost);
            }
            set
            {
                SetStyle(value, swPropMgrPageTextBoxStyle_e.swPropMgrPageTextBoxStyle_NotifyOnlyWhenFocusLost);
            }
        }

        private void Init(int options, String tip = "")
        {
            _swTexteBox = _Groupe.swGroup.AddControl2(Id, (int)swPropertyManagerPageControlType_e.swControlType_Textbox,
                                                                                        tip,
                                                                                        (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge,
                                                                                        options,
                                                                                        tip);
            _swControl = (PropertyManagerPageControl)_swTexteBox;

            _Groupe.Page.DicControl.Add(Id, this);
        }

        public CtrlTextBox(Groupe groupe, int options, String labelTitre = "", String labelTip = "")
            : base(groupe, labelTitre, labelTip)
        {
            Init(options, labelTip);

            _Init = false;
        }

        public CtrlTextBox(Groupe groupe, int options, Boolean avecLabel, Parametre param)
            : base(groupe, param, avecLabel)
        {
            Init(options, param.Tip);

            ApplyParam();

            _Init = false;
        }

        public override void ApplyParam()
        {
            _Init = true;
            if (Param.IsRef())
                TextBoxChanged(this, GetParametre<String>());
            _Init = false;
        }

        public delegate Boolean ValiderTexteEvent(String text);

        public ValiderTexteEvent ValiderTexte;

        public delegate void OnTextBoxChangedEventHandler(Object sender, String text);

        public event OnTextBoxChangedEventHandler OnTextBoxChanged;

        public void TextBoxChanged(Object sender, String text)
        {
            String val = text;

            if (!LectureSeule && (ValiderTexte.IsRef()) && !ValiderTexte(text))
            {
                // On applique l'ancienne valeur
                val = _Text;
                _swTexteBox.Text = val;
            }

            if (_Init || _EcraserTexte)
                _swTexteBox.Text = val;

            _Text = val;
            SetParametre(val);

            if (OnTextBoxChanged.IsRef())
                OnTextBoxChanged(this, text);
        }
    }

    public abstract class CtrlBaseComboBox : Control
    {
        protected PropertyManagerPageCombobox _swComboBox;

        public PropertyManagerPageCombobox swComboBox { get { return _swComboBox; } }

        protected List<String> _Lst = new List<String>();

        private int _SelectedIndex = 0;

        public int SelectedIndex
        {
            get { return _SelectedIndex; }
            set
            {
                if(GetStyle(swPropMgrPageComboBoxStyle_e.swPropMgrPageComboBoxStyle_EditableText))
                    _swComboBox.EditText = _Lst[value];

                SelectionChanged(this, value);
            }
        }

        public short Height { get { return _swComboBox.Height; } set { _swComboBox.Height = value; } }

        protected Boolean GetStyle(swPropMgrPageComboBoxStyle_e p)
        {
            return (_swComboBox.Style & (int)p) > 0;
        }

        protected void SetStyle(Boolean etat, swPropMgrPageComboBoxStyle_e p)
        {
            int s = _swComboBox.Style;
            if (etat)
                _swComboBox.Style = s | (int)p;
            else if (GetStyle(p))
                _swComboBox.Style = ((s - (int)p) > 0) ? (s - (int)p) : s;
        }

        public Boolean Trier
        {
            get
            {
                return GetStyle(swPropMgrPageComboBoxStyle_e.swPropMgrPageComboBoxStyle_Sorted);
            }
            set
            {
                SetStyle(value, swPropMgrPageComboBoxStyle_e.swPropMgrPageComboBoxStyle_Sorted);
            }
        }

        public Boolean NotifieSurSelection
        {
            get
            {
                return GetStyle(swPropMgrPageComboBoxStyle_e.swPropMgrPageComboBoxStyle_AvoidSelectionText);
            }
            set
            {
                SetStyle(value, swPropMgrPageComboBoxStyle_e.swPropMgrPageComboBoxStyle_AvoidSelectionText);
            }
        }

        private void Init(int options, String tip = "")
        {
            _swComboBox = _Groupe.swGroup.AddControl2(Id, (int)swPropertyManagerPageControlType_e.swControlType_Combobox,
                                                                                        tip,
                                                                                        (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge,
                                                                                        options,
                                                                                        tip);
            _swControl = (PropertyManagerPageControl)_swComboBox;

            _Groupe.Page.DicControl.Add(Id, this);
        }

        public CtrlBaseComboBox(Groupe groupe, int options, String labelTitre = "", String labelTip = "")
            : base(groupe, labelTitre, labelTip)
        {
            Init(options, labelTip);

            _Init = false;
        }

        public CtrlBaseComboBox(Groupe groupe, int options, Boolean avecLabel, Parametre param)
            : base(groupe, param, avecLabel)
        {
            Init(options, param.Tip);

            _Init = false;
        }

        public override void ApplyParam()
        {
            _Init = true;
            if (Param.IsRef())
                SelectionChanged(this, OnGetParameter());
            _Init = false;
        }

        public delegate void OnSelectionChangedEventHandler(Object sender, int Item);

        public event OnSelectionChangedEventHandler OnSelectionChanged;

        protected abstract int OnGetParameter();

        protected abstract void OnSetParameter(int index);

        public void SelectionChanged(Object sender, int Index)
        {
            if (_Init)
                _swComboBox.CurrentSelection = (short)Index;

            _SelectedIndex = Index;

            OnSetParameter(Index);

            if (OnSelectionChanged.IsRef())
                OnSelectionChanged(this, Index);
        }

        public List<String> Liste
        {
            get { return _Lst; }
            set
            {
                _swComboBox.Clear();
                _Lst = value;
                _swComboBox.AddItems(_Lst.ToArray());

                if (Param.IsRef())
                    ApplyParam();
                else
                    SelectionChanged(this, 0);
            }
        }

    }

    public class CtrlTextComboBox : CtrlBaseComboBox
    {
        private String _Text = "";

        public String Text
        {
            get { return _Text; }
            set
            {
                if (!LectureSeule)
                    _swComboBox.EditText = value;

                ComboboxEditChanged(this, value);
            }
        }

        public new int SelectedIndex
        {
            get { return base.SelectedIndex; }
            set
            {
                _Text = _Lst[value];
                _swComboBox.CurrentSelection = (short)value;

                SelectionChanged(this, value);
            }
        }

        public T GetTextAs<T>()
        {
            T Val;

            try
            {
                Val = (T)Convert.ChangeType(Text, typeof(T));
            }
            catch
            {
                Val = (T)typeof(T).GetDefaultValue();
            }

            return Val;
        }

        public Boolean LectureSeule
        {
            get
            {
                return GetStyle(swPropMgrPageComboBoxStyle_e.swPropMgrPageComboBoxStyle_EditBoxReadOnly);
            }
            set
            {
                SetStyle(value, swPropMgrPageComboBoxStyle_e.swPropMgrPageComboBoxStyle_EditBoxReadOnly);
            }
        }

        public Boolean Editable
        {
            get
            {
                return GetStyle(swPropMgrPageComboBoxStyle_e.swPropMgrPageComboBoxStyle_EditableText);
            }
            set
            {
                SetStyle(value, swPropMgrPageComboBoxStyle_e.swPropMgrPageComboBoxStyle_EditableText);
            }
        }

        private void Init()
        {
        }

        public CtrlTextComboBox(Groupe groupe, int options, String labelTitre = "", String labelTip = "")
            : base(groupe, options, labelTitre, labelTip)
        {
            Init();

            _Init = false;

            OnSelectionChanged += delegate (Object s, int i) { _Text = _swComboBox.ItemText[(short)i]; };
        }

        public CtrlTextComboBox(Groupe groupe, int options, Boolean avecLabel, Parametre param)
            : base(groupe, options, avecLabel, param)
        {
            Init();
            _Init = false;

            OnSelectionChanged += delegate (Object s, int i) { _Text = _swComboBox.ItemText[(short)i]; };
        }

        public delegate Boolean ValiderTexteEvent(String text);

        public ValiderTexteEvent ValiderTexte;

        public delegate Boolean OnComboboxEditChangedEventHandler(Object sender, String text);

        public event OnComboboxEditChangedEventHandler OnComboboxEditChanged;

        public void ComboboxEditChanged(Object sender, String text)
        {
            String val = text;

            if (!LectureSeule && (ValiderTexte.IsRef()) && !ValiderTexte(text))
                val = "";

            if (_Init && !LectureSeule)
                _swComboBox.EditText = val;

            _Text = val;
            SetParametre(val);

            if (OnComboboxEditChanged.IsRef())
                OnComboboxEditChanged(this, text);
        }

        protected override void OnSetParameter(int index)
        {
            SetParametre(index);
        }

        protected override int OnGetParameter()
        {
            return _Lst.IndexOf(GetParametre<String>());
        }
    }

    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]
    public class CtrlEnumComboBox<T, D> : CtrlBaseComboBox
        where D : PersoAttribute
    {
        private void Init()
        {
            SetStyle(true, swPropMgrPageComboBoxStyle_e.swPropMgrPageComboBoxStyle_EditBoxReadOnly);
            Liste = typeof(T).GetListeEnumInfo<D>();
        }

        public CtrlEnumComboBox(Groupe groupe, int options, Boolean avecLabel, Parametre param)
            : base(groupe, options, avecLabel, param)
        {
            Init();
            _Init = false;
        }

        private int IndexList(T Val)
        {
            return _Lst.IndexOf(Val.GetEnumInfo<D>());
        }

        public T Val
        {
            get { return _Lst[SelectedIndex].GetEnumFromAtt<T, D>(); }
            set { SelectionChanged(this, IndexList(value)); }
        }

        public T FiltrerEnum { set { Liste = value.GetListeEnumInfo<D>(); } }

        protected override void OnSetParameter(int index)
        {
            SetParametre(_Lst[index].GetEnumFromAtt<T, D>());
        }

        protected override int OnGetParameter()
        {
            return IndexList(GetParametre<T>());
        }
    }

    public abstract class CtrlBaseListBox : Control
    {
        protected PropertyManagerPageListbox _swListBox;

        protected List<String> _Lst = new List<String>();

        private int _SelectedIndex = 0;

        public int SelectedIndex
        {
            get { return _SelectedIndex; }
            set
            {
                if ((value < 0) && (value >= _Lst.Count)) return;

                _Init = true;
                SelectionChanged(this, value, true);
                _Init = false;
            }
        }

        private List<int> _ListSelectedIndex = new List<int>();

        public List<int> ListSelectedIndex
        {
            get { return _ListSelectedIndex; }
            set
            {
                for (short i = 0; i < _swListBox.ItemCount; i++)
                {
                    _swListBox.SetSelectedItem(i, false);
                }

                foreach (int index in value)
                    SelectedIndex = index;
            }
        }

        public void ToutSelectionner(Boolean val)
        {
            for (short i = 0; i < _swListBox.ItemCount; i++)
            {
                SelectionChanged(this, i, val);
            }
        }

        public PropertyManagerPageListbox swListBox { get { return _swListBox; } }

        public short Height { get { return _swListBox.Height; } set { _swListBox.Height = value; } }

        private Boolean GetStyle(swPropMgrPageListBoxStyle_e p)
        {
            return (_swListBox.Style & (int)p) > 0;
        }

        private void SetStyle(Boolean etat, swPropMgrPageListBoxStyle_e p)
        {
            int s = _swListBox.Style;
            if (etat)
                _swListBox.Style = s | (int)p;
            else if (GetStyle(p))
                _swListBox.Style = ((s - (int)p) > 0) ? (s - (int)p) : s;
        }

        public Boolean Trier
        {
            get
            {
                return GetStyle(swPropMgrPageListBoxStyle_e.swPropMgrPageListBoxStyle_Sorted);
            }
            set
            {
                SetStyle(value, swPropMgrPageListBoxStyle_e.swPropMgrPageListBoxStyle_Sorted);
            }
        }

        public Boolean SelectionMultiple
        {
            get
            {
                return GetStyle(swPropMgrPageListBoxStyle_e.swPropMgrPageListBoxStyle_MultipleItemSelect);
            }
            set
            {
                SetStyle(value, swPropMgrPageListBoxStyle_e.swPropMgrPageListBoxStyle_MultipleItemSelect);
            }
        }

        public Boolean TouteHauteur
        {
            get
            {
                return GetStyle(swPropMgrPageListBoxStyle_e.swPropMgrPageListBoxStyle_NoIntegralHeight);
            }
            set
            {
                SetStyle(value, swPropMgrPageListBoxStyle_e.swPropMgrPageListBoxStyle_NoIntegralHeight);
            }
        }

        private void Init(int options, String tip = "")
        {
            _swListBox = _Groupe.swGroup.AddControl2(Id, (int)swPropertyManagerPageControlType_e.swControlType_Listbox,
                                                                                        tip,
                                                                                        (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge,
                                                                                        options,
                                                                                        tip);
            _swControl = (PropertyManagerPageControl)_swListBox;

            _Groupe.Page.DicControl.Add(Id, this);
        }

        public CtrlBaseListBox(Groupe groupe, int options, String labelTitre = "", String labelTip = "")
            : base(groupe, labelTitre, labelTip)
        {
            Init(options, labelTip);

            _Init = false;
        }

        public CtrlBaseListBox(Groupe groupe, int options, Boolean avecLabel, Parametre param)
            : base(groupe, param, avecLabel)
        {
            Init(options, param.Tip);

            if (param.IsRef())
                SelectionChanged(this, GetParametre<int>(), true);

            _Init = false;
        }

        public delegate void OnSelectionChangedEventHandler(Object sender, int Item);

        public event OnSelectionChangedEventHandler OnSelectionChanged;

        public void SelectionChanged(Object sender, int Index, Boolean Val = false)
        {
            if (_Init)
                _swListBox.SetSelectedItem((short)Index, Val);

            _SelectedIndex = Index;
            SetParametre(Index);

            _ListSelectedIndex.Clear();

            if (_swListBox.GetSelectedItemsCount() > 0)
            {
                foreach (short index in _swListBox.GetSelectedItems())
                    _ListSelectedIndex.Add((int)index);
            }

            if (ProtectedOnSelectionChanged.IsRef())
                ProtectedOnSelectionChanged(this, Index);

            if (OnSelectionChanged.IsRef())
                OnSelectionChanged(this, Index);
        }

        protected delegate void ProtectedOnSelectionChangedEventHandler(Object sender, int Item);

        protected event ProtectedOnSelectionChangedEventHandler ProtectedOnSelectionChanged;

        public List<String> Liste
        {
            get { return _Lst; }
            set
            {
                _swListBox.Clear();
                _Lst = value;
                _swListBox.AddItems(_Lst.ToArray());

                if (Param.IsRef())
                    SelectionChanged(this, GetParametre<int>(), true);
            }
        }

        public virtual void Vider()
        {
            _swListBox.Clear();
            _Lst.Clear();
            _ListSelectedIndex.Clear();
            _SelectedIndex = 0;
        }
    }

    public class CtrlTextListBox : CtrlBaseListBox
    {
        private String _SelectedText = "";

        private List<String> _ListSelectedText = new List<String>();

        public String SelectedText
        {
            get { return _SelectedText; }
            set
            {
                if (_Lst.Contains(value))
                    SelectedIndex = _Lst.IndexOf(value);
            }
        }

        public List<String> ListSelectedText
        {
            get { return _ListSelectedText; }
        }

        public T GetTextAs<T>()
        {
            T Val;

            try
            {
                Val = (T)Convert.ChangeType(SelectedText, typeof(T));
            }
            catch
            {
                Val = (T)typeof(T).GetDefaultValue();
            }

            return Val;
        }

        public List<T> GetListTextAs<T>()
        {
            List<T> Val = new List<T>();

            foreach (String s in _ListSelectedText)
            {
                try
                {
                    Val.Add((T)Convert.ChangeType(s, typeof(T)));
                }
                catch
                {
                    Val.Add((T)typeof(T).GetDefaultValue());
                }
            }

            return Val;
        }

        private void Init()
        {
            ProtectedOnSelectionChanged += delegate (Object s, int i)
            {
                _SelectedText = _swListBox.ItemText[(short)i];
                _ListSelectedText.Clear();

                if (_swListBox.GetSelectedItemsCount() > 0)
                {
                    foreach (short index in _swListBox.GetSelectedItems())
                        _ListSelectedText.Add(_swListBox.ItemText[index]);
                }
            };
        }

        public CtrlTextListBox(Groupe groupe, int options, String labelTitre = "", String labelTip = "")
            : base(groupe, options, labelTitre, labelTip)
        {
            Init();
        }

        public CtrlTextListBox(Groupe groupe, int options, Boolean avecLabel, Parametre param)
            : base(groupe, options, avecLabel, param)
        {
            Init();
        }

        public override void Vider()
        {
            base.Vider();
            _ListSelectedText.Clear();
            _SelectedText = "";
        }
    }
}
