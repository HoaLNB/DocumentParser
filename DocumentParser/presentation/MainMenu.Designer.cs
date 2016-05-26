namespace DocumentParser.presentation
{
    partial class MainMenu
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
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.txtImgFolderPath = new System.Windows.Forms.TextBox();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btImport = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.btExport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Location = new System.Drawing.Point(12, 80);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(75, 23);
            this.btnSelectFolder.TabIndex = 0;
            this.btnSelectFolder.Text = "Browse";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            this.btnSelectFolder.Click += new System.EventHandler(this.btnSelectFolder_Click);
            // 
            // txtImgFolderPath
            // 
            this.txtImgFolderPath.Location = new System.Drawing.Point(93, 82);
            this.txtImgFolderPath.Name = "txtImgFolderPath";
            this.txtImgFolderPath.Size = new System.Drawing.Size(258, 20);
            this.txtImgFolderPath.TabIndex = 1;
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(12, 30);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(75, 23);
            this.btnSelectFile.TabIndex = 2;
            this.btnSelectFile.Text = "Browse";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(93, 32);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(258, 20);
            this.txtFilePath.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(125, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Select .docx file to import";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 61);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(97, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Select image folder";
            // 
            // btImport
            // 
            this.btImport.Location = new System.Drawing.Point(12, 121);
            this.btImport.Name = "btImport";
            this.btImport.Size = new System.Drawing.Size(107, 38);
            this.btImport.TabIndex = 6;
            this.btImport.Text = "Import";
            this.btImport.UseVisualStyleBackColor = true;
            this.btImport.Click += new System.EventHandler(this.btImport_Click);
            // 
            // btCancel
            // 
            this.btCancel.Location = new System.Drawing.Point(247, 121);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(108, 38);
            this.btCancel.TabIndex = 7;
            this.btCancel.Text = "Cancel";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // btExport
            // 
            this.btExport.Location = new System.Drawing.Point(129, 121);
            this.btExport.Name = "btExport";
            this.btExport.Size = new System.Drawing.Size(112, 38);
            this.btExport.TabIndex = 8;
            this.btExport.Text = "Export";
            this.btExport.UseVisualStyleBackColor = true;
            this.btExport.Click += new System.EventHandler(this.btExport_Click);
            // 
            // MainMenu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(367, 171);
            this.Controls.Add(this.btExport);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.btImport);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.btnSelectFile);
            this.Controls.Add(this.txtImgFolderPath);
            this.Controls.Add(this.btnSelectFolder);
            this.Name = "MainMenu";
            this.Text = "Document Parser 0.1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.TextBox txtImgFolderPath;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btImport;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Button btExport;
    }
}