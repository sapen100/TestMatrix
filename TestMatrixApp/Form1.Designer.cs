namespace TestMatrixApp
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
            this.buttonFolder = new System.Windows.Forms.Button();
            this.labelBulletin = new System.Windows.Forms.Label();
            this.textFolderName = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.logMessage = new System.Windows.Forms.RichTextBox();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.tmButton = new System.Windows.Forms.Button();
            this.folderBrowserDialog2 = new System.Windows.Forms.FolderBrowserDialog();
            this.comboApp = new System.Windows.Forms.ComboBox();
            this.labelComoApp = new System.Windows.Forms.Label();
            this.TApicture = new System.Windows.Forms.PictureBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.resourcesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.bulletinUncompressorToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.configCheckBox = new System.Windows.Forms.CheckBox();
            this.rbLocal = new System.Windows.Forms.RadioButton();
            this.rbWeb = new System.Windows.Forms.RadioButton();
            this.panel1 = new System.Windows.Forms.Panel();
            this.textBulletins = new System.Windows.Forms.TextBox();
            this.labelBulletins = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textLastestTM = new System.Windows.Forms.TextBox();
            this.buttonLatestTM = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.TApicture)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonFolder
            // 
            this.buttonFolder.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.buttonFolder.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.buttonFolder.Location = new System.Drawing.Point(731, 54);
            this.buttonFolder.Name = "buttonFolder";
            this.buttonFolder.Size = new System.Drawing.Size(29, 22);
            this.buttonFolder.TabIndex = 10;
            this.buttonFolder.Text = "...";
            this.buttonFolder.UseVisualStyleBackColor = true;
            this.buttonFolder.Click += new System.EventHandler(this.buttonFolder_Click);
            // 
            // labelBulletin
            // 
            this.labelBulletin.AutoSize = true;
            this.labelBulletin.BackColor = System.Drawing.SystemColors.Window;
            this.labelBulletin.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.labelBulletin.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.labelBulletin.Location = new System.Drawing.Point(38, 58);
            this.labelBulletin.Name = "labelBulletin";
            this.labelBulletin.Size = new System.Drawing.Size(105, 16);
            this.labelBulletin.TabIndex = 9;
            this.labelBulletin.Text = "Bulletin Location";
            this.labelBulletin.Click += new System.EventHandler(this.label2_Click);
            // 
            // textFolderName
            // 
            this.textFolderName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.textFolderName.Location = new System.Drawing.Point(159, 55);
            this.textFolderName.Name = "textFolderName";
            this.textFolderName.Size = new System.Drawing.Size(572, 21);
            this.textFolderName.TabIndex = 8;
            this.textFolderName.Click += new System.EventHandler(this.textFolderName_Click);
            this.textFolderName.TextChanged += new System.EventHandler(this.textFolderName_TextChanged);
            this.textFolderName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textFolderName_KeyDown);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // logMessage
            // 
            this.logMessage.Location = new System.Drawing.Point(4, 81);
            this.logMessage.Name = "logMessage";
            this.logMessage.Size = new System.Drawing.Size(757, 273);
            this.logMessage.TabIndex = 11;
            this.logMessage.Text = "";
            this.logMessage.TextChanged += new System.EventHandler(this.LogMessage_TextChanged);
            // 
            // tmButton
            // 
            this.tmButton.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.tmButton.ImageAlign = System.Drawing.ContentAlignment.TopRight;
            this.tmButton.Location = new System.Drawing.Point(660, 360);
            this.tmButton.Name = "tmButton";
            this.tmButton.Size = new System.Drawing.Size(100, 39);
            this.tmButton.TabIndex = 12;
            this.tmButton.Text = "Start Generating TM";
            this.tmButton.UseVisualStyleBackColor = true;
            this.tmButton.Click += new System.EventHandler(this.button2_Click);
            // 
            // comboApp
            // 
            this.comboApp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboApp.FormattingEnabled = true;
            this.comboApp.Items.AddRange(new object[] {
            "SCAP",
            "FOUNDSTONE",
            "RM"});
            this.comboApp.Location = new System.Drawing.Point(158, 27);
            this.comboApp.Name = "comboApp";
            this.comboApp.Size = new System.Drawing.Size(121, 21);
            this.comboApp.TabIndex = 13;
            this.comboApp.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // labelComoApp
            // 
            this.labelComoApp.AutoSize = true;
            this.labelComoApp.BackColor = System.Drawing.SystemColors.Window;
            this.labelComoApp.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelComoApp.Location = new System.Drawing.Point(45, 34);
            this.labelComoApp.Name = "labelComoApp";
            this.labelComoApp.Size = new System.Drawing.Size(98, 16);
            this.labelComoApp.TabIndex = 14;
            this.labelComoApp.Text = "TM Application";
            // 
            // TApicture
            // 
            this.TApicture.ErrorImage = ((System.Drawing.Image)(resources.GetObject("TApicture.ErrorImage")));
            this.TApicture.Image = ((System.Drawing.Image)(resources.GetObject("TApicture.Image")));
            this.TApicture.Location = new System.Drawing.Point(9, 360);
            this.TApicture.Name = "TApicture";
            this.TApicture.Size = new System.Drawing.Size(139, 38);
            this.TApicture.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.TApicture.TabIndex = 15;
            this.TApicture.TabStop = false;
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.Control;
            this.menuStrip1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.resourcesToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(766, 24);
            this.menuStrip1.TabIndex = 16;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // resourcesToolStripMenuItem
            // 
            this.resourcesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.bulletinUncompressorToolStripMenuItem});
            this.resourcesToolStripMenuItem.Name = "resourcesToolStripMenuItem";
            this.resourcesToolStripMenuItem.Size = new System.Drawing.Size(76, 20);
            this.resourcesToolStripMenuItem.Text = "Resources";
            // 
            // bulletinUncompressorToolStripMenuItem
            // 
            this.bulletinUncompressorToolStripMenuItem.Name = "bulletinUncompressorToolStripMenuItem";
            this.bulletinUncompressorToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.bulletinUncompressorToolStripMenuItem.Text = "Bulletin Uncompressor";
            this.bulletinUncompressorToolStripMenuItem.Click += new System.EventHandler(this.bulletinUncompressorToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(53, 20);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // configCheckBox
            // 
            this.configCheckBox.AutoSize = true;
            this.configCheckBox.Checked = true;
            this.configCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.configCheckBox.Location = new System.Drawing.Point(246, 5);
            this.configCheckBox.Name = "configCheckBox";
            this.configCheckBox.Size = new System.Drawing.Size(100, 17);
            this.configCheckBox.TabIndex = 17;
            this.configCheckBox.Text = "Use SharePoint";
            this.configCheckBox.UseVisualStyleBackColor = true;
            // 
            // rbLocal
            // 
            this.rbLocal.AutoSize = true;
            this.rbLocal.Checked = true;
            this.rbLocal.Location = new System.Drawing.Point(7, 5);
            this.rbLocal.Name = "rbLocal";
            this.rbLocal.Size = new System.Drawing.Size(119, 17);
            this.rbLocal.TabIndex = 18;
            this.rbLocal.TabStop = true;
            this.rbLocal.Text = "From Local Bulletins";
            this.rbLocal.UseVisualStyleBackColor = true;
            this.rbLocal.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // rbWeb
            // 
            this.rbWeb.AutoSize = true;
            this.rbWeb.Location = new System.Drawing.Point(132, 5);
            this.rbWeb.Name = "rbWeb";
            this.rbWeb.Size = new System.Drawing.Size(92, 17);
            this.rbWeb.TabIndex = 19;
            this.rbWeb.Text = "From the Web";
            this.rbWeb.UseVisualStyleBackColor = true;
            this.rbWeb.CheckedChanged += new System.EventHandler(this.rbWeb_CheckedChanged);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.HighlightText;
            this.panel1.Controls.Add(this.rbLocal);
            this.panel1.Controls.Add(this.configCheckBox);
            this.panel1.Controls.Add(this.rbWeb);
            this.panel1.Location = new System.Drawing.Point(413, 27);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(349, 25);
            this.panel1.TabIndex = 20;
            // 
            // textBulletins
            // 
            this.textBulletins.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.textBulletins.Location = new System.Drawing.Point(159, 55);
            this.textBulletins.Name = "textBulletins";
            this.textBulletins.Size = new System.Drawing.Size(601, 21);
            this.textBulletins.TabIndex = 22;
            this.textBulletins.Tag = "hi how are you";
            this.textBulletins.TextChanged += new System.EventHandler(this.textBulletins_TextChanged);
            // 
            // labelBulletins
            // 
            this.labelBulletins.AutoSize = true;
            this.labelBulletins.BackColor = System.Drawing.SystemColors.Window;
            this.labelBulletins.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelBulletins.Location = new System.Drawing.Point(85, 58);
            this.labelBulletins.Name = "labelBulletins";
            this.labelBulletins.Size = new System.Drawing.Size(58, 16);
            this.labelBulletins.TabIndex = 21;
            this.labelBulletins.Text = "Bulletins";
            this.labelBulletins.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.Window;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(13, 83);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(130, 16);
            this.label1.TabIndex = 25;
            this.label1.Text = "TM to Compare With";
            // 
            // textLastestTM
            // 
            this.textLastestTM.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.textLastestTM.Location = new System.Drawing.Point(159, 81);
            this.textLastestTM.Name = "textLastestTM";
            this.textLastestTM.Size = new System.Drawing.Size(572, 21);
            this.textLastestTM.TabIndex = 23;
            // 
            // buttonLatestTM
            // 
            this.buttonLatestTM.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonLatestTM.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.buttonLatestTM.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.buttonLatestTM.Location = new System.Drawing.Point(733, 81);
            this.buttonLatestTM.Name = "buttonLatestTM";
            this.buttonLatestTM.Size = new System.Drawing.Size(28, 21);
            this.buttonLatestTM.TabIndex = 24;
            this.buttonLatestTM.Text = "...";
            this.buttonLatestTM.UseVisualStyleBackColor = true;
            this.buttonLatestTM.Click += new System.EventHandler(this.buttonLatestTM_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(766, 406);
            this.Controls.Add(this.logMessage);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textLastestTM);
            this.Controls.Add(this.buttonLatestTM);
            this.Controls.Add(this.labelBulletins);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.TApicture);
            this.Controls.Add(this.labelComoApp);
            this.Controls.Add(this.comboApp);
            this.Controls.Add(this.tmButton);
            this.Controls.Add(this.labelBulletin);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.buttonFolder);
            this.Controls.Add(this.textFolderName);
            this.Controls.Add(this.textBulletins);
            this.MainMenuStrip = this.menuStrip1;
            this.MaximumSize = new System.Drawing.Size(782, 500);
            this.MinimumSize = new System.Drawing.Size(782, 38);
            this.Name = "Form1";
            this.Text = "Test Matrix Generator V5.7";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.TApicture)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonFolder;
        private System.Windows.Forms.Label labelBulletin;
        private System.Windows.Forms.TextBox textFolderName;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.RichTextBox logMessage;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Button tmButton;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog2;
        private System.Windows.Forms.ComboBox comboApp;
        private System.Windows.Forms.Label labelComoApp;
        private System.Windows.Forms.PictureBox TApicture;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem resourcesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem bulletinUncompressorToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.CheckBox configCheckBox;
        private System.Windows.Forms.RadioButton rbLocal;
        private System.Windows.Forms.RadioButton rbWeb;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox textBulletins;
        private System.Windows.Forms.Label labelBulletins;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textLastestTM;
        private System.Windows.Forms.Button buttonLatestTM;
    }
}

