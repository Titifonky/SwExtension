using LogDebugging;
using ModuleEchelleVue;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Windows.Forms;

namespace SwExtension.ModuleEchelleVue
{
    public partial class FormEchelleVue : Form
    {
        private DrawingDoc Dessin = null;
        private ModelDoc2 Modele = null;
        private Parametre PositionVue = null;
        private ConfigModule Config = null;

        public FormEchelleVue(DrawingDoc dessin, Parametre positionVue, ConfigModule config)
        {
            InitializeComponent();

            Dessin = dessin;
            Modele = dessin.eModelDoc2();
            PositionVue = positionVue;
            Config = config;

            var pos = PositionVue.GetValeur<String>().Split(':');
            System.Drawing.Point pt = new System.Drawing.Point(pos[0].eToInteger(), pos[1].eToInteger());
            Location = pt;

            AjouterEvenement();
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == WM_NCHITTEST)
                m.Result = (IntPtr)(HT_CAPTION);
        }

        private const int WM_NCHITTEST = 0x84;
        private const int HT_CLIENT = 0x1;
        private const int HT_CAPTION = 0x2;

        public void AjouterEvenement()
        {
            Dessin.UserSelectionPostNotify += Dessin_UserSelectionPostNotify;
        }

        public void EnleverEvenement()
        {
            Dessin.UserSelectionPostNotify -= Dessin_UserSelectionPostNotify;
        }

        private int Dessin_UserSelectionPostNotify()
        {
            var typeSel = Modele.eSelect_RecupererSwTypeObjet();

            if(typeSel == swSelectType_e.swSelDRAWINGVIEWS)
            {
                var vue = Modele.eSelect_RecupererObjet<SolidWorks.Interop.sldworks.View>();
                var echelle = (Double[])vue.ScaleRatio;
                TextBoxEchelle.Text = String.Format("{0}:{1}", echelle[0], echelle[1]);
            }

            return 0;
        }

        private void OnClick(object sender, EventArgs e)
        {
            if (Modele.eSelect_Nb() == 0) return;

            var typeSel = Modele.eSelect_RecupererSwTypeObjet();

            if (typeSel == swSelectType_e.swSelDRAWINGVIEWS)
            {
                var vue = Modele.eSelect_RecupererObjet<SolidWorks.Interop.sldworks.View>();
                var echelle = TextBoxEchelle.Text.Split(':');
                if(echelle.Length == 2)
                {
                    var e1 = echelle[0].eToDouble();
                    var e2 = echelle[1].eToDouble();
                    if (e1 > 0 && e2 > 0)
                    {
                        vue.ScaleDecimal = e1 / e2;
                        Modele.EditRebuild3();
                    }
                }
            }
        }

        private void OnClose(object sender, EventArgs e)
        {
            EnleverEvenement();
            Dessin = null;
            var pos = String.Format("{0}:{1}", Location.X, Location.Y);
            PositionVue.SetValeur(pos);
            Config.Sauver();
            this.Close();
            InfoFenetre.Form = null;
        }
    }
}
