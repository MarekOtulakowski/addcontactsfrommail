namespace addcontactsfrommail
{
    partial class F_Main
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(F_Main));
            this.GB_sourceMSFolder = new System.Windows.Forms.GroupBox();
            this.L_folderInfo = new System.Windows.Forms.Label();
            this.L_folder = new System.Windows.Forms.Label();
            this.CB_selectedFolderWithAllSubfolders = new System.Windows.Forms.CheckBox();
            this.RB_sendFolder = new System.Windows.Forms.RadioButton();
            this.B_otherFolder = new System.Windows.Forms.Button();
            this.RB_receiveFolder = new System.Windows.Forms.RadioButton();
            this.L_result = new System.Windows.Forms.Label();
            this.LB_result = new System.Windows.Forms.ListBox();
            this.PB_progress = new System.Windows.Forms.ProgressBar();
            this.B_clearUpperList = new System.Windows.Forms.Button();
            this.B_addContacts = new System.Windows.Forms.Button();
            this.B_abort = new System.Windows.Forms.Button();
            this.BW_worker = new System.ComponentModel.BackgroundWorker();
            this.LL_projectWebsite = new System.Windows.Forms.LinkLabel();
            this.GB_sourceMSFolder.SuspendLayout();
            this.SuspendLayout();
            // 
            // GB_sourceMSFolder
            // 
            this.GB_sourceMSFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.GB_sourceMSFolder.Controls.Add(this.L_folderInfo);
            this.GB_sourceMSFolder.Controls.Add(this.L_folder);
            this.GB_sourceMSFolder.Controls.Add(this.CB_selectedFolderWithAllSubfolders);
            this.GB_sourceMSFolder.Controls.Add(this.RB_sendFolder);
            this.GB_sourceMSFolder.Controls.Add(this.B_otherFolder);
            this.GB_sourceMSFolder.Controls.Add(this.RB_receiveFolder);
            this.GB_sourceMSFolder.Location = new System.Drawing.Point(12, 193);
            this.GB_sourceMSFolder.Name = "GB_sourceMSFolder";
            this.GB_sourceMSFolder.Size = new System.Drawing.Size(307, 106);
            this.GB_sourceMSFolder.TabIndex = 0;
            this.GB_sourceMSFolder.TabStop = false;
            this.GB_sourceMSFolder.Text = "Source MS Outlook folder";
            // 
            // L_folderInfo
            // 
            this.L_folderInfo.AutoSize = true;
            this.L_folderInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.L_folderInfo.Location = new System.Drawing.Point(57, 84);
            this.L_folderInfo.Name = "L_folderInfo";
            this.L_folderInfo.Size = new System.Drawing.Size(27, 13);
            this.L_folderInfo.TabIndex = 10;
            this.L_folderInfo.Text = "null";
            // 
            // L_folder
            // 
            this.L_folder.AutoSize = true;
            this.L_folder.Location = new System.Drawing.Point(16, 84);
            this.L_folder.Name = "L_folder";
            this.L_folder.Size = new System.Drawing.Size(39, 13);
            this.L_folder.TabIndex = 9;
            this.L_folder.Text = "Folder:";
            // 
            // CB_selectedFolderWithAllSubfolders
            // 
            this.CB_selectedFolderWithAllSubfolders.AutoSize = true;
            this.CB_selectedFolderWithAllSubfolders.Location = new System.Drawing.Point(122, 56);
            this.CB_selectedFolderWithAllSubfolders.Name = "CB_selectedFolderWithAllSubfolders";
            this.CB_selectedFolderWithAllSubfolders.Size = new System.Drawing.Size(181, 17);
            this.CB_selectedFolderWithAllSubfolders.TabIndex = 3;
            this.CB_selectedFolderWithAllSubfolders.Text = "selected folder with all subfolders";
            this.CB_selectedFolderWithAllSubfolders.UseVisualStyleBackColor = true;
            // 
            // RB_sendFolder
            // 
            this.RB_sendFolder.AutoSize = true;
            this.RB_sendFolder.Checked = true;
            this.RB_sendFolder.Location = new System.Drawing.Point(17, 29);
            this.RB_sendFolder.Name = "RB_sendFolder";
            this.RB_sendFolder.Size = new System.Drawing.Size(110, 17);
            this.RB_sendFolder.TabIndex = 0;
            this.RB_sendFolder.TabStop = true;
            this.RB_sendFolder.Text = "send mail (default)";
            this.RB_sendFolder.UseVisualStyleBackColor = true;
            this.RB_sendFolder.CheckedChanged += new System.EventHandler(this.RB_sendFolder_CheckedChanged);
            // 
            // B_otherFolder
            // 
            this.B_otherFolder.Location = new System.Drawing.Point(17, 52);
            this.B_otherFolder.Name = "B_otherFolder";
            this.B_otherFolder.Size = new System.Drawing.Size(88, 23);
            this.B_otherFolder.TabIndex = 2;
            this.B_otherFolder.Text = "browse folder";
            this.B_otherFolder.UseVisualStyleBackColor = true;
            this.B_otherFolder.Click += new System.EventHandler(this.B_otherFolder_Click);
            // 
            // RB_receiveFolder
            // 
            this.RB_receiveFolder.AutoSize = true;
            this.RB_receiveFolder.Location = new System.Drawing.Point(142, 29);
            this.RB_receiveFolder.Name = "RB_receiveFolder";
            this.RB_receiveFolder.Size = new System.Drawing.Size(128, 17);
            this.RB_receiveFolder.TabIndex = 1;
            this.RB_receiveFolder.TabStop = true;
            this.RB_receiveFolder.Text = "received mail (default)";
            this.RB_receiveFolder.UseVisualStyleBackColor = true;
            this.RB_receiveFolder.CheckedChanged += new System.EventHandler(this.RB_receiveFolder_CheckedChanged);
            // 
            // L_result
            // 
            this.L_result.AutoSize = true;
            this.L_result.Location = new System.Drawing.Point(12, 8);
            this.L_result.Name = "L_result";
            this.L_result.Size = new System.Drawing.Size(40, 13);
            this.L_result.TabIndex = 13;
            this.L_result.Text = "Result:";
            // 
            // LB_result
            // 
            this.LB_result.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.LB_result.BackColor = System.Drawing.SystemColors.Window;
            this.LB_result.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.LB_result.FormattingEnabled = true;
            this.LB_result.Location = new System.Drawing.Point(12, 27);
            this.LB_result.Name = "LB_result";
            this.LB_result.Size = new System.Drawing.Size(308, 160);
            this.LB_result.TabIndex = 4;
            // 
            // PB_progress
            // 
            this.PB_progress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.PB_progress.Location = new System.Drawing.Point(12, 311);
            this.PB_progress.Name = "PB_progress";
            this.PB_progress.Size = new System.Drawing.Size(307, 23);
            this.PB_progress.TabIndex = 11;
            // 
            // B_clearUpperList
            // 
            this.B_clearUpperList.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.B_clearUpperList.Location = new System.Drawing.Point(227, 340);
            this.B_clearUpperList.Name = "B_clearUpperList";
            this.B_clearUpperList.Size = new System.Drawing.Size(92, 23);
            this.B_clearUpperList.TabIndex = 3;
            this.B_clearUpperList.Text = "Clear result list";
            this.B_clearUpperList.UseVisualStyleBackColor = true;
            this.B_clearUpperList.Click += new System.EventHandler(this.B_clearUpperList_Click);
            // 
            // B_addContacts
            // 
            this.B_addContacts.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.B_addContacts.Location = new System.Drawing.Point(12, 340);
            this.B_addContacts.Name = "B_addContacts";
            this.B_addContacts.Size = new System.Drawing.Size(92, 23);
            this.B_addContacts.TabIndex = 1;
            this.B_addContacts.Text = "Add contacts";
            this.B_addContacts.UseVisualStyleBackColor = true;
            this.B_addContacts.Click += new System.EventHandler(this.B_addContacts_Click);
            // 
            // B_abort
            // 
            this.B_abort.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.B_abort.Location = new System.Drawing.Point(120, 340);
            this.B_abort.Name = "B_abort";
            this.B_abort.Size = new System.Drawing.Size(92, 23);
            this.B_abort.TabIndex = 2;
            this.B_abort.Text = "Abort";
            this.B_abort.UseVisualStyleBackColor = true;
            this.B_abort.Click += new System.EventHandler(this.B_abort_Click);
            // 
            // BW_worker
            // 
            this.BW_worker.WorkerReportsProgress = true;
            this.BW_worker.WorkerSupportsCancellation = true;
            // 
            // LL_projectWebsite
            // 
            this.LL_projectWebsite.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.LL_projectWebsite.AutoSize = true;
            this.LL_projectWebsite.Location = new System.Drawing.Point(243, 369);
            this.LL_projectWebsite.Name = "LL_projectWebsite";
            this.LL_projectWebsite.Size = new System.Drawing.Size(78, 13);
            this.LL_projectWebsite.TabIndex = 14;
            this.LL_projectWebsite.TabStop = true;
            this.LL_projectWebsite.Text = "project website";
            this.LL_projectWebsite.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LL_projectWebsite_LinkClicked);
            // 
            // F_Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(332, 388);
            this.Controls.Add(this.LL_projectWebsite);
            this.Controls.Add(this.B_abort);
            this.Controls.Add(this.GB_sourceMSFolder);
            this.Controls.Add(this.L_result);
            this.Controls.Add(this.LB_result);
            this.Controls.Add(this.PB_progress);
            this.Controls.Add(this.B_clearUpperList);
            this.Controls.Add(this.B_addContacts);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "F_Main";
            this.Text = "Add Contacts From Mail v.1.0.0.5";
            this.Load += new System.EventHandler(this.F_Main_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.F_Main_FormClosing);
            this.GB_sourceMSFolder.ResumeLayout(false);
            this.GB_sourceMSFolder.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox GB_sourceMSFolder;
        private System.Windows.Forms.CheckBox CB_selectedFolderWithAllSubfolders;
        private System.Windows.Forms.RadioButton RB_sendFolder;
        private System.Windows.Forms.Button B_otherFolder;
        private System.Windows.Forms.RadioButton RB_receiveFolder;
        private System.Windows.Forms.Label L_result;
        private System.Windows.Forms.ListBox LB_result;
        private System.Windows.Forms.ProgressBar PB_progress;
        private System.Windows.Forms.Button B_clearUpperList;
        private System.Windows.Forms.Button B_addContacts;
        private System.Windows.Forms.Label L_folderInfo;
        private System.Windows.Forms.Label L_folder;
        private System.Windows.Forms.Button B_abort;
        private System.Windows.Forms.LinkLabel LL_projectWebsite;


    }
}

