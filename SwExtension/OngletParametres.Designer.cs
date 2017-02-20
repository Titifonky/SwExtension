namespace SwExtension
{
    partial class OngletParametres
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
            this.Importer = new System.Windows.Forms.Button();
            this.ListBoxParams = new System.Windows.Forms.ListBox();
            this.Composant = new System.Windows.Forms.TextBox();
            this.MaJ = new System.Windows.Forms.Button();
            this.EditParametre = new System.Windows.Forms.TextBox();
            this.ValidParam = new System.Windows.Forms.Button();
            this.radioMdlCourant = new System.Windows.Forms.RadioButton();
            this.radioTtModeles = new System.Windows.Forms.RadioButton();
            this.ValidProp = new System.Windows.Forms.Button();
            this.EditPropriete = new System.Windows.Forms.TextBox();
            this.ListBoxPropMdl = new System.Windows.Forms.ListBox();
            this.ListBoxPropCfg = new System.Windows.Forms.ListBox();
            this.Groupe1 = new System.Windows.Forms.GroupBox();
            this.radioCfg = new System.Windows.Forms.RadioButton();
            this.radioMdl = new System.Windows.Forms.RadioButton();
            this.Groupe2 = new System.Windows.Forms.GroupBox();
            this.Scanner = new System.Windows.Forms.Button();
            this.Groupe1.SuspendLayout();
            this.Groupe2.SuspendLayout();
            this.SuspendLayout();
            // 
            // Importer
            // 
            this.Importer.Location = new System.Drawing.Point(5, 134);
            this.Importer.Name = "Importer";
            this.Importer.Size = new System.Drawing.Size(267, 25);
            this.Importer.TabIndex = 0;
            this.Importer.Text = "Importer";
            this.Importer.UseVisualStyleBackColor = true;
            this.Importer.Click += new System.EventHandler(this.Importer_Click);
            // 
            // ListBoxParams
            // 
            this.ListBoxParams.AllowDrop = true;
            this.ListBoxParams.FormattingEnabled = true;
            this.ListBoxParams.Location = new System.Drawing.Point(5, 231);
            this.ListBoxParams.Name = "ListBoxParams";
            this.ListBoxParams.Size = new System.Drawing.Size(267, 160);
            this.ListBoxParams.TabIndex = 3;
            this.ListBoxParams.SelectedIndexChanged += new System.EventHandler(this.ListeBoxParams_SelectionChanged);
            this.ListBoxParams.DragDrop += new System.Windows.Forms.DragEventHandler(this.ListBoxParams_DragDrop);
            this.ListBoxParams.DragOver += new System.Windows.Forms.DragEventHandler(this.ListBoxParams_DragOver);
            this.ListBoxParams.MouseDown += new System.Windows.Forms.MouseEventHandler(this.ListBoxParams_MouseDown);
            // 
            // Composant
            // 
            this.Composant.Location = new System.Drawing.Point(5, 5);
            this.Composant.Name = "Composant";
            this.Composant.ReadOnly = true;
            this.Composant.Size = new System.Drawing.Size(267, 20);
            this.Composant.TabIndex = 4;
            // 
            // MaJ
            // 
            this.MaJ.Location = new System.Drawing.Point(5, 169);
            this.MaJ.Name = "MaJ";
            this.MaJ.Size = new System.Drawing.Size(267, 25);
            this.MaJ.TabIndex = 5;
            this.MaJ.Text = "Mettre à jour";
            this.MaJ.UseVisualStyleBackColor = true;
            this.MaJ.Click += new System.EventHandler(this.Maj_Click);
            // 
            // EditParametre
            // 
            this.EditParametre.Location = new System.Drawing.Point(5, 204);
            this.EditParametre.Name = "EditParametre";
            this.EditParametre.Size = new System.Drawing.Size(222, 20);
            this.EditParametre.TabIndex = 7;
            // 
            // ValidParam
            // 
            this.ValidParam.Location = new System.Drawing.Point(233, 202);
            this.ValidParam.Name = "ValidParam";
            this.ValidParam.Size = new System.Drawing.Size(39, 23);
            this.ValidParam.TabIndex = 8;
            this.ValidParam.Text = "Ok";
            this.ValidParam.UseVisualStyleBackColor = true;
            this.ValidParam.Click += new System.EventHandler(this.ValidParam_Click);
            // 
            // radioMdlCourant
            // 
            this.radioMdlCourant.AutoSize = true;
            this.radioMdlCourant.Checked = true;
            this.radioMdlCourant.Location = new System.Drawing.Point(5, 19);
            this.radioMdlCourant.Name = "radioMdlCourant";
            this.radioMdlCourant.Size = new System.Drawing.Size(81, 17);
            this.radioMdlCourant.TabIndex = 11;
            this.radioMdlCourant.TabStop = true;
            this.radioMdlCourant.Text = "Mdl courant";
            this.radioMdlCourant.UseVisualStyleBackColor = true;
            // 
            // radioTtModeles
            // 
            this.radioTtModeles.AutoSize = true;
            this.radioTtModeles.Location = new System.Drawing.Point(5, 42);
            this.radioTtModeles.Name = "radioTtModeles";
            this.radioTtModeles.Size = new System.Drawing.Size(93, 17);
            this.radioTtModeles.TabIndex = 12;
            this.radioTtModeles.Text = "Tt les modeles";
            this.radioTtModeles.UseVisualStyleBackColor = true;
            // 
            // ValidProp
            // 
            this.ValidProp.Location = new System.Drawing.Point(233, 432);
            this.ValidProp.Name = "ValidProp";
            this.ValidProp.Size = new System.Drawing.Size(39, 23);
            this.ValidProp.TabIndex = 14;
            this.ValidProp.Text = "Ok";
            this.ValidProp.UseVisualStyleBackColor = true;
            this.ValidProp.Click += new System.EventHandler(this.ValiProp_Click);
            // 
            // EditPropriete
            // 
            this.EditPropriete.Location = new System.Drawing.Point(5, 434);
            this.EditPropriete.Name = "EditPropriete";
            this.EditPropriete.Size = new System.Drawing.Size(222, 20);
            this.EditPropriete.TabIndex = 13;
            // 
            // ListBoxPropMdl
            // 
            this.ListBoxPropMdl.AllowDrop = true;
            this.ListBoxPropMdl.FormattingEnabled = true;
            this.ListBoxPropMdl.Location = new System.Drawing.Point(5, 461);
            this.ListBoxPropMdl.Name = "ListBoxPropMdl";
            this.ListBoxPropMdl.Size = new System.Drawing.Size(267, 186);
            this.ListBoxPropMdl.TabIndex = 15;
            this.ListBoxPropMdl.SelectedIndexChanged += new System.EventHandler(this.ListBoxProp_SelectionChanged);
            // 
            // ListBoxPropCfg
            // 
            this.ListBoxPropCfg.AllowDrop = true;
            this.ListBoxPropCfg.FormattingEnabled = true;
            this.ListBoxPropCfg.Location = new System.Drawing.Point(5, 661);
            this.ListBoxPropCfg.Name = "ListBoxPropCfg";
            this.ListBoxPropCfg.Size = new System.Drawing.Size(267, 108);
            this.ListBoxPropCfg.TabIndex = 16;
            this.ListBoxPropCfg.SelectedIndexChanged += new System.EventHandler(this.ListBoxProp_SelectionChanged);
            // 
            // Groupe1
            // 
            this.Groupe1.Controls.Add(this.radioCfg);
            this.Groupe1.Controls.Add(this.radioMdl);
            this.Groupe1.Location = new System.Drawing.Point(5, 31);
            this.Groupe1.Name = "Groupe1";
            this.Groupe1.Size = new System.Drawing.Size(80, 67);
            this.Groupe1.TabIndex = 17;
            this.Groupe1.TabStop = false;
            this.Groupe1.Text = "Appliquer à";
            // 
            // radioCfg
            // 
            this.radioCfg.AutoSize = true;
            this.radioCfg.Location = new System.Drawing.Point(6, 42);
            this.radioCfg.Name = "radioCfg";
            this.radioCfg.Size = new System.Drawing.Size(55, 17);
            this.radioCfg.TabIndex = 11;
            this.radioCfg.Text = "Config";
            this.radioCfg.UseVisualStyleBackColor = true;
            // 
            // radioMdl
            // 
            this.radioMdl.AutoSize = true;
            this.radioMdl.Checked = true;
            this.radioMdl.Location = new System.Drawing.Point(6, 19);
            this.radioMdl.Name = "radioMdl";
            this.radioMdl.Size = new System.Drawing.Size(60, 17);
            this.radioMdl.TabIndex = 10;
            this.radioMdl.TabStop = true;
            this.radioMdl.Text = "Modele";
            this.radioMdl.UseVisualStyleBackColor = true;
            // 
            // Groupe2
            // 
            this.Groupe2.Controls.Add(this.radioMdlCourant);
            this.Groupe2.Controls.Add(this.radioTtModeles);
            this.Groupe2.Location = new System.Drawing.Point(91, 31);
            this.Groupe2.Name = "Groupe2";
            this.Groupe2.Size = new System.Drawing.Size(113, 67);
            this.Groupe2.TabIndex = 18;
            this.Groupe2.TabStop = false;
            this.Groupe2.Text = "Appliquer à";
            // 
            // Scanner
            // 
            this.Scanner.Location = new System.Drawing.Point(5, 103);
            this.Scanner.Name = "Scanner";
            this.Scanner.Size = new System.Drawing.Size(267, 25);
            this.Scanner.TabIndex = 19;
            this.Scanner.Text = "Scanner et ajouter aux param";
            this.Scanner.UseVisualStyleBackColor = true;
            this.Scanner.Click += new System.EventHandler(this.Scanner_Click);
            // 
            // OngletParametres
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 772);
            this.ControlBox = false;
            this.Controls.Add(this.Scanner);
            this.Controls.Add(this.Groupe2);
            this.Controls.Add(this.Groupe1);
            this.Controls.Add(this.ListBoxPropCfg);
            this.Controls.Add(this.ListBoxPropMdl);
            this.Controls.Add(this.ValidProp);
            this.Controls.Add(this.EditPropriete);
            this.Controls.Add(this.ValidParam);
            this.Controls.Add(this.EditParametre);
            this.Controls.Add(this.MaJ);
            this.Controls.Add(this.Composant);
            this.Controls.Add(this.ListBoxParams);
            this.Controls.Add(this.Importer);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "OngletParametres";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Resize += new System.EventHandler(this.OnResize);
            this.Groupe1.ResumeLayout(false);
            this.Groupe1.PerformLayout();
            this.Groupe2.ResumeLayout(false);
            this.Groupe2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Importer;
        private System.Windows.Forms.ListBox ListBoxParams;
        private System.Windows.Forms.TextBox Composant;
        private System.Windows.Forms.Button MaJ;
        private System.Windows.Forms.TextBox EditParametre;
        private System.Windows.Forms.Button ValidParam;
        private System.Windows.Forms.RadioButton radioMdlCourant;
        private System.Windows.Forms.RadioButton radioTtModeles;
        private System.Windows.Forms.Button ValidProp;
        private System.Windows.Forms.TextBox EditPropriete;
        private System.Windows.Forms.ListBox ListBoxPropMdl;
        private System.Windows.Forms.ListBox ListBoxPropCfg;
        private System.Windows.Forms.GroupBox Groupe1;
        private System.Windows.Forms.RadioButton radioCfg;
        private System.Windows.Forms.RadioButton radioMdl;
        private System.Windows.Forms.GroupBox Groupe2;
        private System.Windows.Forms.Button Scanner;
    }
}