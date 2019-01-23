namespace SwExtension.ModuleEchelleVue
{
    partial class FormEchelleVue
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
            this.TextBoxEchelle = new System.Windows.Forms.TextBox();
            this.Appliquer = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // TextBoxEchelle
            // 
            this.TextBoxEchelle.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TextBoxEchelle.Location = new System.Drawing.Point(12, 7);
            this.TextBoxEchelle.Name = "TextBoxEchelle";
            this.TextBoxEchelle.Size = new System.Drawing.Size(101, 21);
            this.TextBoxEchelle.TabIndex = 0;
            // 
            // Appliquer
            // 
            this.Appliquer.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Appliquer.Location = new System.Drawing.Point(119, 6);
            this.Appliquer.Name = "Appliquer";
            this.Appliquer.Size = new System.Drawing.Size(75, 23);
            this.Appliquer.TabIndex = 1;
            this.Appliquer.Text = "Appliquer";
            this.Appliquer.UseVisualStyleBackColor = true;
            this.Appliquer.Click += new System.EventHandler(this.OnClick);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(200, 6);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(24, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "X";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.OnClose);
            // 
            // FormEchelleVue
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightGray;
            this.ClientSize = new System.Drawing.Size(230, 34);
            this.ControlBox = false;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.Appliquer);
            this.Controls.Add(this.TextBoxEchelle);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormEchelleVue";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TextBoxEchelle;
        private System.Windows.Forms.Button Appliquer;
        private System.Windows.Forms.Button button1;
    }
}