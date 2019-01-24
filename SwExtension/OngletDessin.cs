using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SwExtension
{
    public partial class OngletDessin : Form
    {
        private SldWorks _Sw = null;
        private ModelDoc2 MdlActif = null;
        private DrawingDoc DessinActif = null;

        public int b = 5;
        public int j = 3;

        public OngletDessin(SldWorks sw)
        {
            InitializeComponent();

            _Sw = sw;
        }

        private void OnResize(object sender, EventArgs e) { ResizeControl(); }

        private void ResizeControl()
        {
            //int width = ClientRectangle.Width - 2 * b;
            //LabelFeuille.Location = new Point(LabelFeuille.Location.X, b);

            //TextBoxFeuille.Location = new Point(TextBoxFeuille.Location.X, LabelFeuille.Location.Y + LabelFeuille.Height + j);
            //TextBoxFeuille.Width = width;

            //AppliquerFeuille.Location = new Point(AppliquerFeuille.Location.X, TextBoxFeuille.Location.Y + TextBoxFeuille.Height + j);

            //LabelVue.Location = new Point(LabelVue.Location.X, AppliquerFeuille.Location.Y + AppliquerFeuille.Height + (3 * j));

            //TextBoxVue.Location = new Point(TextBoxVue.Location.X, LabelVue.Location.Y + LabelVue.Height + j);
            //TextBoxVue.Width = width;

            //AppliquerVue.Location = new Point(AppliquerVue.Location.X, TextBoxVue.Location.Y + TextBoxVue.Height + j);
        }

        public int ActiveDocChange()
        {
            MdlActif = _Sw.ActiveDoc;

            if (DessinActif.IsNull() && MdlActif.IsRef() && (MdlActif.TypeDoc() == eTypeDoc.Dessin))
            {
                DessinActif = MdlActif.eDrawingDoc();
                AjouterEvenement();
                DessinActif_ActivateSheetPostNotify(DessinActif.eFeuilleActive().GetName());
            }
            else if (MdlActif.IsRef() && (MdlActif.TypeDoc() != eTypeDoc.Dessin))
            {
                Deconnecter();
            }

            return 1;
        }

        public void Deconnecter()
        {
            ReinitialiserFeuille();
            EnleverEvenement();
            MdlActif = null;
            DessinActif = null;
        }

        public int CloseDoc(String nomFichier, int raison)
        {
            ReinitialiserFeuille();
            EnleverEvenement();
            MdlActif = null;
            DessinActif = null;

            return 1;
        }

        public void AjouterEvenement()
        {
            DessinActif.ActivateSheetPostNotify += DessinActif_ActivateSheetPostNotify;
            DessinActif.UserSelectionPostNotify += Dessin_UserSelectionPostNotify;
        }

        public void EnleverEvenement()
        {
            DessinActif.ActivateSheetPostNotify -= DessinActif_ActivateSheetPostNotify;
            DessinActif.UserSelectionPostNotify -= Dessin_UserSelectionPostNotify;
        }

        private int DessinActif_ActivateSheetPostNotify(string SheetName)
        {
            LabelFeuille.Text = String.Format("Feuille : {0}", SheetName);
            var Feuille = DessinActif.Sheet[SheetName];
            var prop = (Double[])Feuille.GetProperties2();
            TextBoxFeuille.Text = String.Format("{0}:{1}", prop[2], prop[3]);
            ReinitialiserVue();
            return 0;
        }

        private Boolean InitTextBoxVue = false;

        private int Dessin_UserSelectionPostNotify()
        {
            var typeSel = MdlActif.eSelect_RecupererSwTypeObjet();

            if (typeSel == swSelectType_e.swSelDRAWINGVIEWS)
            {
                InitTextBoxVue = true;

                var vue = MdlActif.eSelect_RecupererObjet<SolidWorks.Interop.sldworks.View>();

                LabelVue.Text = String.Format("Vue : {0}", vue.GetName2());

                var echelle = (Double[])vue.ScaleRatio;
                var vueParent = (SolidWorks.Interop.sldworks.View)vue.GetBaseView();

                if (vueParent.IsNull())
                {
                    BtParent.Checked = false;
                    BtParent.Enabled = false;
                }
                else
                    BtParent.Enabled = true;

                if (vueParent.IsRef() && vue.UseParentScale)
                    BtParent.Checked = true;
                else if (vue.UseSheetScale.ToBoolean())
                    BtFeuille.Checked = true;
                else
                    BtPersonnalise.Checked = true;

                TextBoxVue.Text = String.Format("{0}:{1}", echelle[0], echelle[1]);

                InitTextBoxVue = false;
            }
            else
            {
                if (LabelVue.Text != "Vue")
                    ReinitialiserVue();
            }

            return 0;
        }

        private void ReinitialiserFeuille()
        {
            LabelFeuille.Text = "Feuille";
            TextBoxFeuille.Text = "";
            ReinitialiserVue();
        }

        private void ReinitialiserVue()
        {
            LabelVue.Text = "Vue";
            TextBoxVue.Text = "";
            BtFeuille.Checked = false;
            BtParent.Checked = false;
            BtParent.Enabled = true;
            BtPersonnalise.Checked = false;
            CurrentClickVue = false;
        }

        private void OnClickFeuille(object sender, EventArgs e)
        {
            try
            {
                if (DessinActif.IsNull()) return;

                var Feuille = DessinActif.eFeuilleActive();
                var echelle = TextBoxFeuille.Text.Split(':');
                if (echelle.Length == 2)
                {
                    var e1 = echelle[0].eToDouble();
                    var e2 = echelle[1].eToDouble();
                    if (e1 > 0 && e2 > 0)
                    {
                        Feuille.SetScale(e1, e2, false, false);
                        MdlActif.EditRebuild3();
                    }
                }
            }
            catch
            {
                Deconnecter();
            }
        }

        private Boolean CurrentClickVue = false;

        private void OnClickVue(object sender, EventArgs e)
        {
            if (CurrentClickVue) return;

            try
            {
                if (DessinActif.IsNull() || (MdlActif.eSelect_Nb() == 0)) return;

                var typeSel = MdlActif.eSelect_RecupererSwTypeObjet();

                if (typeSel == swSelectType_e.swSelDRAWINGVIEWS)
                {
                    CurrentClickVue = true;

                    var vue = MdlActif.eSelect_RecupererObjet<SolidWorks.Interop.sldworks.View>();

                    if (BtParent.Checked)
                        vue.UseParentScale = true;
                    else if (BtFeuille.Checked)
                        vue.UseSheetScale = 1;
                    else if (BtPersonnalise.Checked)
                    {
                        var echelle = TextBoxVue.Text.Split(':');
                        if (echelle.Length == 2)
                        {
                            var e1 = echelle[0].eToDouble();
                            var e2 = echelle[1].eToDouble();
                            if (e1 > 0 && e2 > 0)
                            {
                                vue.ScaleRatio = new Double[] { e1, e2 };
                                //vue.ScaleDecimal = e1 / e2;
                                MdlActif.ForceRebuild3(true);
                                MdlActif.EditRebuild3();
                                MdlActif.GraphicsRedraw2();
                            }
                        }
                    }

                    vue.eSelectionner(DessinActif);

                    CurrentClickVue = false;
                }
            }
            catch
            {
                ReinitialiserVue();
            }
        }

        private void TextBoxVue_TextChanged(object sender, EventArgs e)
        {
            if (!InitTextBoxVue)
                BtPersonnalise.Checked = true;
        }
    }
}
