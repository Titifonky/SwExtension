namespace SwExtension
{
    partial class OngletLog
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
            this.Afficher_Log = new System.Windows.Forms.Button();
            this.TexteLog = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // Afficher_Log
            // 
            this.Afficher_Log.Location = new System.Drawing.Point(5, 5);
            this.Afficher_Log.Name = "Afficher_Log";
            this.Afficher_Log.Size = new System.Drawing.Size(267, 25);
            this.Afficher_Log.TabIndex = 0;
            this.Afficher_Log.Text = "Afficher Log";
            this.Afficher_Log.UseVisualStyleBackColor = true;
            this.Afficher_Log.Click += new System.EventHandler(this.OnClick);
            // 
            // TexteLog
            // 
            this.TexteLog.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TexteLog.Location = new System.Drawing.Point(5, 37);
            this.TexteLog.Name = "TexteLog";
            this.TexteLog.Size = new System.Drawing.Size(267, 376);
            this.TexteLog.TabIndex = 2;
            this.TexteLog.Text = "";
            // 
            // OngletLog_New
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 425);
            this.ControlBox = false;
            this.Controls.Add(this.TexteLog);
            this.Controls.Add(this.Afficher_Log);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "OngletLog_New";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Resize += new System.EventHandler(this.OnResize);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Afficher_Log;
        private System.Windows.Forms.RichTextBox TexteLog;
    }
}