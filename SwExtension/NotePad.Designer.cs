namespace SwExtension
{
    partial class NotePad
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
            this.TexteLog = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // TexteLog
            // 
            this.TexteLog.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TexteLog.Location = new System.Drawing.Point(-3, 0);
            this.TexteLog.Name = "TexteLog";
            this.TexteLog.Size = new System.Drawing.Size(387, 467);
            this.TexteLog.TabIndex = 1;
            this.TexteLog.Text = "";
            // 
            // NotePad
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CausesValidation = false;
            this.ClientSize = new System.Drawing.Size(383, 468);
            this.Controls.Add(this.TexteLog);
            this.Name = "NotePad";
            this.ShowIcon = false;
            this.Text = "Log";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.OnClose);
            this.Resize += new System.EventHandler(this.OnResize);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.RichTextBox TexteLog;
    }
}