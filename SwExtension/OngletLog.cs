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
    public partial class OngletLog : Form
    {
        public int b = 5;
        public int j = 3;

        public OngletLog()
        {
            InitializeComponent();

            Afficher_Log.Padding = new Padding(0);
            Afficher_Log.Margin = new Padding(0);

            TexteLog.Multiline = true;
            TexteLog.ReadOnly = true;
            TexteLog.BackColor = Color.LightGray;
            TexteLog.WordWrap = false;
            TexteLog.ScrollBars = RichTextBoxScrollBars.Both;
            TexteLog.Padding = new Padding(0);
            TexteLog.Margin = new Padding(0);
            ResizeControl();
        }

        private void OnResize(object sender, EventArgs e) { ResizeControl(); }

        private void ResizeControl()
        {
            Afficher_Log.Location = new Point(b, b);
            Afficher_Log.Size = new Size(ClientRectangle.Width - 2 * b, Afficher_Log.Height);

            TexteLog.Location = new Point(b, b + Afficher_Log.Height + j);
            TexteLog.Size = new Size(Afficher_Log.Width, ClientRectangle.Height - (b + Afficher_Log.Height + j + b));
        }

        public RichTextBox Texte
        {
            get { return TexteLog; }
        }

        private void OnClick(object sender, EventArgs e)
        {
            LogDebugging.WindowLog.AfficherFenetre(true);
        }
    }
}
