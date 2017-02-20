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
    public partial class NotePad : Form
    {
        public NotePad()
        {
            InitializeComponent();
            TexteLog.Multiline = true;
            TexteLog.ReadOnly = true;
            TexteLog.BackColor = Color.LightGray;
            TexteLog.BorderStyle = BorderStyle.None;
            TexteLog.WordWrap = false;
            TexteLog.ScrollBars = RichTextBoxScrollBars.Both;
            TexteLog.Padding = new Padding(4);
            TexteLog.Margin = new Padding(4);
            ResizeControl();
        }

        private void OnResize(object sender, EventArgs e) { ResizeControl(); }

        private void ResizeControl()
        {
            TexteLog.Location = new Point(0, 0);
            TexteLog.Size = new Size(ClientRectangle.Width, ClientRectangle.Height);
        }

        public void AppendText(String t)
        {
            TexteLog.AppendText(t);
            TexteLog.SelectionStart = TexteLog.TextLength;
            TexteLog.ScrollToCaret();
            this.Refresh();
        }

        public void Effacer()
        {
            TexteLog.Text = "";
        }

        private void OnClose(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }
    }
}
