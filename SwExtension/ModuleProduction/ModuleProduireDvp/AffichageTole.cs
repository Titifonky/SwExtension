using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ModuleProduction.ModuleProduireDvp
{
    public partial class AffichageTole : Form
    {
        public AffichageTole()
        {
            InitializeComponent();

            Resize += delegate { ListeView.Size = new Size(this.Size.Width - 40, this.Size.Height - 70); };

            //ListeView.Dock = DockStyle.Fill;
            ListeView.View = View.Tile;

            ListeView.TileSize = new Size(200, 200);

        }

        private void ValiderForm(object sender, MouseEventArgs e)
        {
            this.Close();
        }

        private Boolean mouseDown;
        private Point lastLocation;

        private void AffichageTole_MouseDown(object sender, MouseEventArgs e)
        {
            mouseDown = true;
            lastLocation = e.Location;
        }

        private void AffichageTole_MouseMove(object sender, MouseEventArgs e)
        {
            if(mouseDown)
            {
                this.Location = new Point((this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);
                this.Update();
            }
        }

        private void AffichageTole_MouseUp(object sender, MouseEventArgs e)
        {
            mouseDown = false;
        }
    }
}
