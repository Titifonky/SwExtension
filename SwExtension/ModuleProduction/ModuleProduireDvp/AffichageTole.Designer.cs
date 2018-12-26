namespace ModuleProduction.ModuleProduireDvp
{
    partial class AffichageTole
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.ListeView = new System.Windows.Forms.ListView();
            this.Valider = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ListeView
            // 
            this.ListeView.Location = new System.Drawing.Point(12, 42);
            this.ListeView.Name = "ListeView";
            this.ListeView.Size = new System.Drawing.Size(946, 347);
            this.ListeView.TabIndex = 0;
            this.ListeView.UseCompatibleStateImageBehavior = false;
            // 
            // Valider
            // 
            this.Valider.Location = new System.Drawing.Point(13, 13);
            this.Valider.Name = "Valider";
            this.Valider.Size = new System.Drawing.Size(75, 23);
            this.Valider.TabIndex = 1;
            this.Valider.Text = "Ok";
            this.Valider.UseVisualStyleBackColor = true;
            this.Valider.MouseClick += new System.Windows.Forms.MouseEventHandler(this.ValiderForm);
            // 
            // AffichageTole
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(970, 401);
            this.ControlBox = false;
            this.Controls.Add(this.Valider);
            this.Controls.Add(this.ListeView);
            this.Name = "AffichageTole";
            this.Padding = new System.Windows.Forms.Padding(5, 30, 5, 5);
            this.ShowIcon = false;
            this.TopMost = true;
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.AffichageTole_MouseDown);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.AffichageTole_MouseMove);
            this.MouseUp += new System.Windows.Forms.MouseEventHandler(this.AffichageTole_MouseUp);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListView ListeView;
        private System.Windows.Forms.Button Valider;
    }
}