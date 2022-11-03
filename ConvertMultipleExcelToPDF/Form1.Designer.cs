namespace ConvertMultipleExcelToPDF
{
    partial class FrmMain
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMain));
            this.TxtDraggedFiles = new System.Windows.Forms.Label();
            this.BtnLoad = new System.Windows.Forms.Button();
            this.labelDragFolder = new System.Windows.Forms.Label();
            this.labelInfo = new System.Windows.Forms.Label();
            this.BtnLetsGo = new System.Windows.Forms.Button();
            this.labelErrorMessage = new System.Windows.Forms.Label();
            this.IconError = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.labelFooter1 = new System.Windows.Forms.Label();
            this.labelFooter2 = new System.Windows.Forms.Label();
            this.labelDescription = new System.Windows.Forms.Label();
            this.checkBoxAllWorkBook = new System.Windows.Forms.CheckBox();
            this.checkBoxDragFiles = new System.Windows.Forms.CheckBox();
            this.toolTipDrag = new System.Windows.Forms.ToolTip(this.components);
            this.ImgGit = new System.Windows.Forms.PictureBox();
            this.pictureLogo = new System.Windows.Forms.PictureBox();
            this.pictureDrag = new System.Windows.Forms.PictureBox();
            this.LoadingImage = new System.Windows.Forms.PictureBox();
            this.TxtBoxLoad = new System.Windows.Forms.TextBox();
            this.picDone = new System.Windows.Forms.PictureBox();
            this.LabelEmptyXLS = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.ImgGit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureLogo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureDrag)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImage)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picDone)).BeginInit();
            this.SuspendLayout();
            // 
            // TxtDraggedFiles
            // 
            this.TxtDraggedFiles.AutoSize = true;
            this.TxtDraggedFiles.ForeColor = System.Drawing.SystemColors.Highlight;
            this.TxtDraggedFiles.Location = new System.Drawing.Point(6, 13);
            this.TxtDraggedFiles.Name = "TxtDraggedFiles";
            this.TxtDraggedFiles.Size = new System.Drawing.Size(142, 16);
            this.TxtDraggedFiles.TabIndex = 0;
            this.TxtDraggedFiles.Text = "Drag your Excel files ...";
            this.TxtDraggedFiles.Visible = false;
            // 
            // BtnLoad
            // 
            this.BtnLoad.BackColor = System.Drawing.Color.OrangeRed;
            this.BtnLoad.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BtnLoad.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.BtnLoad.FlatAppearance.MouseDownBackColor = System.Drawing.Color.OrangeRed;
            this.BtnLoad.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(95)))), ((int)(((byte)(34)))));
            this.BtnLoad.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BtnLoad.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.BtnLoad.Location = new System.Drawing.Point(389, 6);
            this.BtnLoad.Name = "BtnLoad";
            this.BtnLoad.Size = new System.Drawing.Size(182, 31);
            this.BtnLoad.TabIndex = 0;
            this.BtnLoad.TabStop = false;
            this.BtnLoad.Text = "Load ...";
            this.BtnLoad.UseVisualStyleBackColor = false;
            this.BtnLoad.Click += new System.EventHandler(this.BtnLoad_Click);
            // 
            // labelDragFolder
            // 
            this.labelDragFolder.AutoSize = true;
            this.labelDragFolder.ForeColor = System.Drawing.Color.DimGray;
            this.labelDragFolder.Location = new System.Drawing.Point(391, 41);
            this.labelDragFolder.Name = "labelDragFolder";
            this.labelDragFolder.Size = new System.Drawing.Size(179, 16);
            this.labelDragFolder.TabIndex = 0;
            this.labelDragFolder.Text = "Click here or drag your folder";
            // 
            // labelInfo
            // 
            this.labelInfo.AutoSize = true;
            this.labelInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelInfo.ForeColor = System.Drawing.Color.ForestGreen;
            this.labelInfo.Location = new System.Drawing.Point(6, 36);
            this.labelInfo.Name = "labelInfo";
            this.labelInfo.Size = new System.Drawing.Size(17, 16);
            this.labelInfo.TabIndex = 0;
            this.labelInfo.Text = "...";
            // 
            // BtnLetsGo
            // 
            this.BtnLetsGo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BtnLetsGo.Location = new System.Drawing.Point(66, 58);
            this.BtnLetsGo.Name = "BtnLetsGo";
            this.BtnLetsGo.Size = new System.Drawing.Size(239, 34);
            this.BtnLetsGo.TabIndex = 0;
            this.BtnLetsGo.TabStop = false;
            this.BtnLetsGo.Text = "Let\'s Go";
            this.BtnLetsGo.UseVisualStyleBackColor = true;
            this.BtnLetsGo.Click += new System.EventHandler(this.BtnLetsGo_Click);
            // 
            // labelErrorMessage
            // 
            this.labelErrorMessage.AutoSize = true;
            this.labelErrorMessage.BackColor = System.Drawing.SystemColors.Control;
            this.labelErrorMessage.ForeColor = System.Drawing.Color.Red;
            this.labelErrorMessage.Location = new System.Drawing.Point(28, 102);
            this.labelErrorMessage.Name = "labelErrorMessage";
            this.labelErrorMessage.Size = new System.Drawing.Size(0, 16);
            this.labelErrorMessage.TabIndex = 0;
            // 
            // IconError
            // 
            this.IconError.AutoSize = true;
            this.IconError.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.IconError.ForeColor = System.Drawing.Color.Red;
            this.IconError.Location = new System.Drawing.Point(6, 98);
            this.IconError.Name = "IconError";
            this.IconError.Size = new System.Drawing.Size(16, 24);
            this.IconError.TabIndex = 0;
            this.IconError.Text = "!";
            this.IconError.Visible = false;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Silver;
            this.button1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.DarkGray;
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Gainsboro;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(388, 20);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(182, 31);
            this.button1.TabIndex = 0;
            this.button1.TabStop = false;
            this.button1.Text = "...";
            this.button1.UseVisualStyleBackColor = false;
            // 
            // labelFooter1
            // 
            this.labelFooter1.AutoSize = true;
            this.labelFooter1.BackColor = System.Drawing.SystemColors.Control;
            this.labelFooter1.ForeColor = System.Drawing.Color.DimGray;
            this.labelFooter1.Location = new System.Drawing.Point(4, 220);
            this.labelFooter1.Name = "labelFooter1";
            this.labelFooter1.Size = new System.Drawing.Size(130, 16);
            this.labelFooter1.TabIndex = 4;
            this.labelFooter1.Text = "© 2022 Excel to PDF.";
            // 
            // labelFooter2
            // 
            this.labelFooter2.AutoSize = true;
            this.labelFooter2.BackColor = System.Drawing.SystemColors.Control;
            this.labelFooter2.ForeColor = System.Drawing.Color.DimGray;
            this.labelFooter2.Location = new System.Drawing.Point(442, 220);
            this.labelFooter2.Name = "labelFooter2";
            this.labelFooter2.Size = new System.Drawing.Size(135, 16);
            this.labelFooter2.TabIndex = 5;
            this.labelFooter2.Text = "Abdessalam AADEL.";
            // 
            // labelDescription
            // 
            this.labelDescription.AutoSize = true;
            this.labelDescription.BackColor = System.Drawing.Color.Transparent;
            this.labelDescription.ForeColor = System.Drawing.Color.DarkSlateGray;
            this.labelDescription.Location = new System.Drawing.Point(6, 131);
            this.labelDescription.Name = "labelDescription";
            this.labelDescription.Size = new System.Drawing.Size(232, 64);
            this.labelDescription.TabIndex = 0;
            this.labelDescription.Text = "Convert Multiple Excel file: is a simple \r\nprograme that allow you can convert \r\n" +
    "Multiple Excel files in multiple Folder \r\nto PDF in same Location.";
            // 
            // checkBoxAllWorkBook
            // 
            this.checkBoxAllWorkBook.AutoSize = true;
            this.checkBoxAllWorkBook.BackColor = System.Drawing.Color.Transparent;
            this.checkBoxAllWorkBook.Cursor = System.Windows.Forms.Cursors.Hand;
            this.checkBoxAllWorkBook.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxAllWorkBook.ForeColor = System.Drawing.Color.OrangeRed;
            this.checkBoxAllWorkBook.Location = new System.Drawing.Point(311, 57);
            this.checkBoxAllWorkBook.Name = "checkBoxAllWorkBook";
            this.checkBoxAllWorkBook.Size = new System.Drawing.Size(120, 20);
            this.checkBoxAllWorkBook.TabIndex = 7;
            this.checkBoxAllWorkBook.Text = "Entire workbook";
            this.toolTipDrag.SetToolTip(this.checkBoxAllWorkBook, "Activate this option to convert All sheet to PDF,\r\nif was not checked, it only co" +
        "nverts Active sheet(s).");
            this.checkBoxAllWorkBook.UseVisualStyleBackColor = false;
            this.checkBoxAllWorkBook.CheckedChanged += new System.EventHandler(this.checkBoxAllWorkBook_CheckedChanged);
            // 
            // checkBoxDragFiles
            // 
            this.checkBoxDragFiles.AutoSize = true;
            this.checkBoxDragFiles.BackColor = System.Drawing.Color.Transparent;
            this.checkBoxDragFiles.Cursor = System.Windows.Forms.Cursors.Hand;
            this.checkBoxDragFiles.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxDragFiles.ForeColor = System.Drawing.Color.OrangeRed;
            this.checkBoxDragFiles.Location = new System.Drawing.Point(311, 73);
            this.checkBoxDragFiles.Name = "checkBoxDragFiles";
            this.checkBoxDragFiles.Size = new System.Drawing.Size(117, 20);
            this.checkBoxDragFiles.TabIndex = 8;
            this.checkBoxDragFiles.Text = "Drag Excel files";
            this.toolTipDrag.SetToolTip(this.checkBoxDragFiles, "Activate this option to Drag & Drop\r\njust Excel files.");
            this.checkBoxDragFiles.UseVisualStyleBackColor = false;
            this.checkBoxDragFiles.CheckedChanged += new System.EventHandler(this.checkBoxDragFiles_CheckedChanged);
            // 
            // ImgGit
            // 
            this.ImgGit.Cursor = System.Windows.Forms.Cursors.Hand;
            this.ImgGit.Image = ((System.Drawing.Image)(resources.GetObject("ImgGit.Image")));
            this.ImgGit.Location = new System.Drawing.Point(291, 214);
            this.ImgGit.Name = "ImgGit";
            this.ImgGit.Size = new System.Drawing.Size(20, 20);
            this.ImgGit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.ImgGit.TabIndex = 10;
            this.ImgGit.TabStop = false;
            this.toolTipDrag.SetToolTip(this.ImgGit, "Go to Github Repository.");
            this.ImgGit.Click += new System.EventHandler(this.ImgGit_Click);
            // 
            // pictureLogo
            // 
            this.pictureLogo.Image = ((System.Drawing.Image)(resources.GetObject("pictureLogo.Image")));
            this.pictureLogo.Location = new System.Drawing.Point(390, 63);
            this.pictureLogo.Name = "pictureLogo";
            this.pictureLogo.Size = new System.Drawing.Size(180, 142);
            this.pictureLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureLogo.TabIndex = 6;
            this.pictureLogo.TabStop = false;
            // 
            // pictureDrag
            // 
            this.pictureDrag.Image = ((System.Drawing.Image)(resources.GetObject("pictureDrag.Image")));
            this.pictureDrag.Location = new System.Drawing.Point(250, 87);
            this.pictureDrag.Name = "pictureDrag";
            this.pictureDrag.Size = new System.Drawing.Size(152, 127);
            this.pictureDrag.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureDrag.TabIndex = 3;
            this.pictureDrag.TabStop = false;
            this.pictureDrag.Visible = false;
            // 
            // LoadingImage
            // 
            this.LoadingImage.BackColor = System.Drawing.Color.Transparent;
            this.LoadingImage.Image = global::ConvertMultipleExcelToPDF.Properties.Resources.loading;
            this.LoadingImage.Location = new System.Drawing.Point(77, 1);
            this.LoadingImage.Name = "LoadingImage";
            this.LoadingImage.Size = new System.Drawing.Size(161, 50);
            this.LoadingImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.LoadingImage.TabIndex = 9;
            this.LoadingImage.TabStop = false;
            this.LoadingImage.Visible = false;
            // 
            // TxtBoxLoad
            // 
            this.TxtBoxLoad.ForeColor = System.Drawing.SystemColors.GrayText;
            this.TxtBoxLoad.Location = new System.Drawing.Point(8, 11);
            this.TxtBoxLoad.Name = "TxtBoxLoad";
            this.TxtBoxLoad.ReadOnly = true;
            this.TxtBoxLoad.Size = new System.Drawing.Size(374, 22);
            this.TxtBoxLoad.TabIndex = 11;
            this.TxtBoxLoad.Text = "Chose your folder location ...";
            // 
            // picDone
            // 
            this.picDone.Image = ((System.Drawing.Image)(resources.GetObject("picDone.Image")));
            this.picDone.Location = new System.Drawing.Point(47, 38);
            this.picDone.Name = "picDone";
            this.picDone.Size = new System.Drawing.Size(14, 14);
            this.picDone.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picDone.TabIndex = 12;
            this.picDone.TabStop = false;
            this.picDone.Visible = false;
            // 
            // LabelEmptyXLS
            // 
            this.LabelEmptyXLS.AutoSize = true;
            this.LabelEmptyXLS.BackColor = System.Drawing.Color.Transparent;
            this.LabelEmptyXLS.Cursor = System.Windows.Forms.Cursors.Hand;
            this.LabelEmptyXLS.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LabelEmptyXLS.ForeColor = System.Drawing.Color.DodgerBlue;
            this.LabelEmptyXLS.Location = new System.Drawing.Point(163, 37);
            this.LabelEmptyXLS.Name = "LabelEmptyXLS";
            this.LabelEmptyXLS.Size = new System.Drawing.Size(82, 16);
            this.LabelEmptyXLS.TabIndex = 13;
            this.LabelEmptyXLS.Text = "Empty Excel";
            this.LabelEmptyXLS.Visible = false;
            this.LabelEmptyXLS.Click += new System.EventHandler(this.LabelEmptyXLS_Click);
            // 
            // FrmMain
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(580, 238);
            this.Controls.Add(this.LabelEmptyXLS);
            this.Controls.Add(this.picDone);
            this.Controls.Add(this.TxtBoxLoad);
            this.Controls.Add(this.ImgGit);
            this.Controls.Add(this.checkBoxDragFiles);
            this.Controls.Add(this.checkBoxAllWorkBook);
            this.Controls.Add(this.labelDescription);
            this.Controls.Add(this.labelFooter2);
            this.Controls.Add(this.labelFooter1);
            this.Controls.Add(this.IconError);
            this.Controls.Add(this.labelErrorMessage);
            this.Controls.Add(this.BtnLetsGo);
            this.Controls.Add(this.labelInfo);
            this.Controls.Add(this.labelDragFolder);
            this.Controls.Add(this.BtnLoad);
            this.Controls.Add(this.TxtDraggedFiles);
            this.Controls.Add(this.pictureLogo);
            this.Controls.Add(this.pictureDrag);
            this.Controls.Add(this.LoadingImage);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "FrmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Convert Multiple Excel File to PDF";
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.FrmMain_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.FrmMain_DragEnter);
            this.DragLeave += new System.EventHandler(this.FrmMain_DragLeave);
            ((System.ComponentModel.ISupportInitialize)(this.ImgGit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureLogo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureDrag)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImage)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picDone)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label TxtDraggedFiles;
        private System.Windows.Forms.Button BtnLoad;
        private System.Windows.Forms.Label labelDragFolder;
        private System.Windows.Forms.PictureBox pictureDrag;
        private System.Windows.Forms.Label labelInfo;
        private System.Windows.Forms.Button BtnLetsGo;
        private System.Windows.Forms.Label labelErrorMessage;
        private System.Windows.Forms.Label IconError;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label labelFooter1;
        private System.Windows.Forms.Label labelFooter2;
        private System.Windows.Forms.PictureBox pictureLogo;
        private System.Windows.Forms.Label labelDescription;
        private System.Windows.Forms.CheckBox checkBoxAllWorkBook;
        private System.Windows.Forms.CheckBox checkBoxDragFiles;
        private System.Windows.Forms.PictureBox LoadingImage;
        private System.Windows.Forms.ToolTip toolTipDrag;
        private System.Windows.Forms.PictureBox ImgGit;
        private System.Windows.Forms.TextBox TxtBoxLoad;
        private System.Windows.Forms.PictureBox picDone;
        private System.Windows.Forms.Label LabelEmptyXLS;
    }
}

