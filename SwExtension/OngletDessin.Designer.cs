namespace SwExtension
{
    partial class OngletDessin
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
            this.LabelFeuille = new System.Windows.Forms.Label();
            this.TextBoxFeuille = new System.Windows.Forms.TextBox();
            this.LabelVue = new System.Windows.Forms.Label();
            this.TextBoxVue = new System.Windows.Forms.TextBox();
            this.AppliquerFeuille = new System.Windows.Forms.Button();
            this.AppliquerVue = new System.Windows.Forms.Button();
            this.BtFeuille = new System.Windows.Forms.RadioButton();
            this.BtParent = new System.Windows.Forms.RadioButton();
            this.BtPersonnalise = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // LabelFeuille
            // 
            this.LabelFeuille.AutoSize = true;
            this.LabelFeuille.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LabelFeuille.Location = new System.Drawing.Point(2, 9);
            this.LabelFeuille.Name = "LabelFeuille";
            this.LabelFeuille.Size = new System.Drawing.Size(51, 15);
            this.LabelFeuille.TabIndex = 0;
            this.LabelFeuille.Text = "Feuille";
            // 
            // TextBoxFeuille
            // 
            this.TextBoxFeuille.Location = new System.Drawing.Point(5, 27);
            this.TextBoxFeuille.Name = "TextBoxFeuille";
            this.TextBoxFeuille.Size = new System.Drawing.Size(116, 20);
            this.TextBoxFeuille.TabIndex = 1;
            // 
            // LabelVue
            // 
            this.LabelVue.AutoSize = true;
            this.LabelVue.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LabelVue.Location = new System.Drawing.Point(2, 98);
            this.LabelVue.Name = "LabelVue";
            this.LabelVue.Size = new System.Drawing.Size(31, 15);
            this.LabelVue.TabIndex = 2;
            this.LabelVue.Text = "Vue";
            // 
            // TextBoxVue
            // 
            this.TextBoxVue.Location = new System.Drawing.Point(5, 188);
            this.TextBoxVue.Name = "TextBoxVue";
            this.TextBoxVue.Size = new System.Drawing.Size(116, 20);
            this.TextBoxVue.TabIndex = 3;
            this.TextBoxVue.TextChanged += new System.EventHandler(this.TextBoxVue_TextChanged);
            // 
            // AppliquerFeuille
            // 
            this.AppliquerFeuille.Location = new System.Drawing.Point(5, 53);
            this.AppliquerFeuille.Name = "AppliquerFeuille";
            this.AppliquerFeuille.Size = new System.Drawing.Size(75, 23);
            this.AppliquerFeuille.TabIndex = 4;
            this.AppliquerFeuille.Text = "Appliquer";
            this.AppliquerFeuille.UseVisualStyleBackColor = true;
            this.AppliquerFeuille.Click += new System.EventHandler(this.OnClickFeuille);
            // 
            // AppliquerVue
            // 
            this.AppliquerVue.Location = new System.Drawing.Point(5, 214);
            this.AppliquerVue.Name = "AppliquerVue";
            this.AppliquerVue.Size = new System.Drawing.Size(75, 23);
            this.AppliquerVue.TabIndex = 5;
            this.AppliquerVue.Text = "Appliquer";
            this.AppliquerVue.UseVisualStyleBackColor = true;
            this.AppliquerVue.Click += new System.EventHandler(this.OnClickVue);
            // 
            // BtFeuille
            // 
            this.BtFeuille.AutoSize = true;
            this.BtFeuille.Location = new System.Drawing.Point(5, 117);
            this.BtFeuille.Name = "BtFeuille";
            this.BtFeuille.Size = new System.Drawing.Size(116, 17);
            this.BtFeuille.TabIndex = 6;
            this.BtFeuille.TabStop = true;
            this.BtFeuille.Text = "Echelle de la feuille";
            this.BtFeuille.UseVisualStyleBackColor = true;
            // 
            // BtParent
            // 
            this.BtParent.AutoSize = true;
            this.BtParent.Location = new System.Drawing.Point(5, 141);
            this.BtParent.Name = "BtParent";
            this.BtParent.Size = new System.Drawing.Size(108, 17);
            this.BtParent.TabIndex = 7;
            this.BtParent.TabStop = true;
            this.BtParent.Text = "Echelle du parent";
            this.BtParent.UseVisualStyleBackColor = true;
            // 
            // BtPersonnalise
            // 
            this.BtPersonnalise.AutoSize = true;
            this.BtPersonnalise.Location = new System.Drawing.Point(5, 165);
            this.BtPersonnalise.Name = "BtPersonnalise";
            this.BtPersonnalise.Size = new System.Drawing.Size(91, 17);
            this.BtPersonnalise.TabIndex = 8;
            this.BtPersonnalise.TabStop = true;
            this.BtPersonnalise.Text = "Personnalisé :";
            this.BtPersonnalise.UseVisualStyleBackColor = true;
            // 
            // OngletDessin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(268, 386);
            this.ControlBox = false;
            this.Controls.Add(this.BtPersonnalise);
            this.Controls.Add(this.BtParent);
            this.Controls.Add(this.BtFeuille);
            this.Controls.Add(this.AppliquerVue);
            this.Controls.Add(this.AppliquerFeuille);
            this.Controls.Add(this.TextBoxVue);
            this.Controls.Add(this.LabelVue);
            this.Controls.Add(this.TextBoxFeuille);
            this.Controls.Add(this.LabelFeuille);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "OngletDessin";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "OngletDessin";
            this.Resize += new System.EventHandler(this.OnResize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label LabelFeuille;
        private System.Windows.Forms.TextBox TextBoxFeuille;
        private System.Windows.Forms.Label LabelVue;
        private System.Windows.Forms.TextBox TextBoxVue;
        private System.Windows.Forms.Button AppliquerFeuille;
        private System.Windows.Forms.Button AppliquerVue;
        private System.Windows.Forms.RadioButton BtFeuille;
        private System.Windows.Forms.RadioButton BtParent;
        private System.Windows.Forms.RadioButton BtPersonnalise;
    }
}