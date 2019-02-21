namespace ScoreSheets
{
    partial class View
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
            this.SelectFile = new System.Windows.Forms.Button();
            this.RegionalCheck = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // SelectFile
            // 
            this.SelectFile.Location = new System.Drawing.Point(303, 59);
            this.SelectFile.Name = "SelectFile";
            this.SelectFile.Size = new System.Drawing.Size(254, 199);
            this.SelectFile.TabIndex = 0;
            this.SelectFile.Text = "Signups";
            this.SelectFile.UseVisualStyleBackColor = true;
            this.SelectFile.Click += new System.EventHandler(this.SelectFile_Click);
            // 
            // RegionalCheck
            // 
            this.RegionalCheck.Location = new System.Drawing.Point(29, 59);
            this.RegionalCheck.Name = "RegionalCheck";
            this.RegionalCheck.Size = new System.Drawing.Size(247, 199);
            this.RegionalCheck.TabIndex = 1;
            this.RegionalCheck.Text = "For Regionals Signups";
            this.RegionalCheck.UseVisualStyleBackColor = true;
            this.RegionalCheck.MouseClick += new System.Windows.Forms.MouseEventHandler(this.RegionalCheck_MouseClick);
            // 
            // View
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(581, 326);
            this.Controls.Add(this.RegionalCheck);
            this.Controls.Add(this.SelectFile);
            this.Name = "View";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Model_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button SelectFile;
        private System.Windows.Forms.Button RegionalCheck;
    }
}

