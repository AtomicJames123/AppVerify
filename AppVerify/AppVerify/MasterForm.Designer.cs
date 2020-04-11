namespace AppVerify
{
    partial class MasterForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MasterForm));
            this.RunChecklist = new System.Windows.Forms.Button();
            this.StudentVerify = new System.Windows.Forms.Button();
            this.F = new System.Windows.Forms.Button();
            this.RemoveFiles = new System.Windows.Forms.Button();
            this.InstalledApps = new System.Windows.Forms.Button();
            this.About = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // RunChecklist
            // 
            this.RunChecklist.Location = new System.Drawing.Point(192, 38);
            this.RunChecklist.Margin = new System.Windows.Forms.Padding(2);
            this.RunChecklist.Name = "RunChecklist";
            this.RunChecklist.Size = new System.Drawing.Size(134, 26);
            this.RunChecklist.TabIndex = 0;
            this.RunChecklist.Text = "Run CheckList";
            this.RunChecklist.UseVisualStyleBackColor = true;
            this.RunChecklist.Click += new System.EventHandler(this.RunChecklist_Click);
            // 
            // StudentVerify
            // 
            this.StudentVerify.Location = new System.Drawing.Point(192, 92);
            this.StudentVerify.Margin = new System.Windows.Forms.Padding(2);
            this.StudentVerify.Name = "StudentVerify";
            this.StudentVerify.Size = new System.Drawing.Size(134, 26);
            this.StudentVerify.TabIndex = 1;
            this.StudentVerify.Text = "Student App Verification";
            this.StudentVerify.UseVisualStyleBackColor = true;
            this.StudentVerify.Click += new System.EventHandler(this.StudentVerify_Click);
            // 
            // F
            // 
            this.F.Location = new System.Drawing.Point(192, 144);
            this.F.Margin = new System.Windows.Forms.Padding(2);
            this.F.Name = "F";
            this.F.Size = new System.Drawing.Size(134, 26);
            this.F.TabIndex = 2;
            this.F.Text = "Teacher App Verification";
            this.F.UseVisualStyleBackColor = true;
            this.F.Click += new System.EventHandler(this.TeacherVerify_Click);
            // 
            // RemoveFiles
            // 
            this.RemoveFiles.Location = new System.Drawing.Point(414, 38);
            this.RemoveFiles.Margin = new System.Windows.Forms.Padding(2);
            this.RemoveFiles.Name = "RemoveFiles";
            this.RemoveFiles.Size = new System.Drawing.Size(88, 42);
            this.RemoveFiles.TabIndex = 3;
            this.RemoveFiles.Text = "Remove AppVerify Files";
            this.RemoveFiles.UseVisualStyleBackColor = true;
            this.RemoveFiles.Click += new System.EventHandler(this.RemovedFiles_Click);
            // 
            // InstalledApps
            // 
            this.InstalledApps.Location = new System.Drawing.Point(192, 196);
            this.InstalledApps.Margin = new System.Windows.Forms.Padding(2);
            this.InstalledApps.Name = "InstalledApps";
            this.InstalledApps.Size = new System.Drawing.Size(134, 26);
            this.InstalledApps.TabIndex = 4;
            this.InstalledApps.Text = "Installed Applications";
            this.InstalledApps.UseVisualStyleBackColor = true;
            this.InstalledApps.Click += new System.EventHandler(this.InstalledApps_Click);
            // 
            // About
            // 
            this.About.Location = new System.Drawing.Point(432, 253);
            this.About.Margin = new System.Windows.Forms.Padding(2);
            this.About.Name = "About";
            this.About.Size = new System.Drawing.Size(70, 25);
            this.About.TabIndex = 5;
            this.About.Text = "About";
            this.About.UseVisualStyleBackColor = true;
            this.About.Click += new System.EventHandler(this.About_Click);
            // 
            // MasterForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(521, 292);
            this.Controls.Add(this.About);
            this.Controls.Add(this.InstalledApps);
            this.Controls.Add(this.RemoveFiles);
            this.Controls.Add(this.F);
            this.Controls.Add(this.StudentVerify);
            this.Controls.Add(this.RunChecklist);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MaximizeBox = false;
            this.Name = "MasterForm";
            this.Text = "AppVerify";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button RunChecklist;
        private System.Windows.Forms.Button StudentVerify;
        private System.Windows.Forms.Button F;
        private System.Windows.Forms.Button RemoveFiles;
        private System.Windows.Forms.Button InstalledApps;
        private System.Windows.Forms.Button About;
    }
}

