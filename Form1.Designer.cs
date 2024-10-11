namespace test
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.txtPDFPath = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.lectureName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.filePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCut = new System.Windows.Forms.Button();
            this.btnChonFile = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // txtPDFPath
            // 
            resources.ApplyResources(this.txtPDFPath, "txtPDFPath");
            this.txtPDFPath.Name = "txtPDFPath";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // fileSystemWatcher1
            // 
            this.fileSystemWatcher1.EnableRaisingEvents = true;
            this.fileSystemWatcher1.SynchronizingObject = this;
            // 
            // dataGridView1
            // 
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.lectureName,
            this.filePath});
            resources.ApplyResources(this.dataGridView1, "dataGridView1");
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            // 
            // lectureName
            // 
            resources.ApplyResources(this.lectureName, "lectureName");
            this.lectureName.Name = "lectureName";
            // 
            // filePath
            // 
            resources.ApplyResources(this.filePath, "filePath");
            this.filePath.Name = "filePath";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Image = global::test.Properties.Resources.link;
            this.label1.Name = "label1";
            // 
            // btnCut
            // 
            this.btnCut.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnCut.BackgroundImage = global::test.Properties.Resources.export__1_;
            resources.ApplyResources(this.btnCut, "btnCut");
            this.btnCut.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnCut.FlatAppearance.BorderSize = 0;
            this.btnCut.Name = "btnCut";
            this.btnCut.UseVisualStyleBackColor = false;
            this.btnCut.Click += new System.EventHandler(this.btnCut_Click);
            // 
            // btnChonFile
            // 
            this.btnChonFile.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnChonFile.BackgroundImage = global::test.Properties.Resources.folder__1_;
            resources.ApplyResources(this.btnChonFile, "btnChonFile");
            this.btnChonFile.ForeColor = System.Drawing.Color.Black;
            this.btnChonFile.Name = "btnChonFile";
            this.btnChonFile.UseVisualStyleBackColor = false;
            this.btnChonFile.Click += new System.EventHandler(this.btnChonFile_Click);
            // 
            // Form1
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnCut);
            this.Controls.Add(this.txtPDFPath);
            this.Controls.Add(this.btnChonFile);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnChonFile;
        private System.Windows.Forms.TextBox txtPDFPath;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.Windows.Forms.Button btnCut;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn lectureName;
        private System.Windows.Forms.DataGridViewTextBoxColumn filePath;
        private System.Windows.Forms.Label label1;
    }
}

