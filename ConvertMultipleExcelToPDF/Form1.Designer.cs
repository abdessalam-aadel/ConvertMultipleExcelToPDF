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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMain));
            this.TxtFolderName = new System.Windows.Forms.Label();
            this.BtnLoad = new System.Windows.Forms.Button();
            this.labelDragFolder = new System.Windows.Forms.Label();
            this.pictureDrag = new System.Windows.Forms.PictureBox();
            this.labelInfo = new System.Windows.Forms.Label();
            this.BtnLetsGo = new System.Windows.Forms.Button();
            this.labelErrorMessage = new System.Windows.Forms.Label();
            this.IconError = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.labelFooter1 = new System.Windows.Forms.Label();
            this.labelFooter2 = new System.Windows.Forms.Label();
            this.pictureLogo = new System.Windows.Forms.PictureBox();
            this.labelDescription = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureDrag)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureLogo)).BeginInit();
            this.SuspendLayout();
            // 
            // TxtFolderName
            // 
            this.TxtFolderName.AutoSize = true;
            this.TxtFolderName.ForeColor = System.Drawing.SystemColors.Highlight;
            this.TxtFolderName.Location = new System.Drawing.Point(6, 13);
            this.TxtFolderName.Name = "TxtFolderName";
            this.TxtFolderName.Size = new System.Drawing.Size(184, 16);
            this.TxtFolderName.TabIndex = 0;
            this.TxtFolderName.Text = "Chose your Folder Location ...";
            // 
            // BtnLoad
            // 
            this.BtnLoad.BackColor = System.Drawing.Color.OrangeRed;
            this.BtnLoad.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BtnLoad.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.BtnLoad.FlatAppearance.MouseDownBackColor = System.Drawing.Color.OrangeRed;
            this.BtnLoad.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(95)))), ((int)(((byte)(34)))));
            this.BtnLoad.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BtnLoad.ForeColor = System.Drawing.Color.White;
            this.BtnLoad.Location = new System.Drawing.Point(389, 6);
            this.BtnLoad.Name = "BtnLoad";
            this.BtnLoad.Size = new System.Drawing.Size(182, 31);
            this.BtnLoad.TabIndex = 0;
            this.BtnLoad.TabStop = false;
            this.BtnLoad.Text = "...";
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
            // labelInfo
            // 
            this.labelInfo.AutoSize = true;
            this.labelInfo.ForeColor = System.Drawing.Color.ForestGreen;
            this.labelInfo.Location = new System.Drawing.Point(6, 35);
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
            this.labelFooter1.Location = new System.Drawing.Point(6, 201);
            this.labelFooter1.Name = "labelFooter1";
            this.labelFooter1.Size = new System.Drawing.Size(130, 16);
            this.labelFooter1.TabIndex = 4;
            this.labelFooter1.Text = "© 2020 Excel to PDF.";
            // 
            // labelFooter2
            // 
            this.labelFooter2.AutoSize = true;
            this.labelFooter2.BackColor = System.Drawing.SystemColors.Control;
            this.labelFooter2.ForeColor = System.Drawing.Color.DimGray;
            this.labelFooter2.Location = new System.Drawing.Point(444, 201);
            this.labelFooter2.Name = "labelFooter2";
            this.labelFooter2.Size = new System.Drawing.Size(135, 16);
            this.labelFooter2.TabIndex = 5;
            this.labelFooter2.Text = "Abdessalam AADEL.";
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
            // labelDescription
            // 
            this.labelDescription.AutoSize = true;
            this.labelDescription.BackColor = System.Drawing.Color.Transparent;
            this.labelDescription.ForeColor = System.Drawing.Color.DarkSlateGray;
            this.labelDescription.Location = new System.Drawing.Point(6, 131);
            this.labelDescription.Name = "labelDescription";
            this.labelDescription.Size = new System.Drawing.Size(237, 64);
            this.labelDescription.TabIndex = 0;
            this.labelDescription.Text = "Convert Multiple Excel file: is a simple \r\nprograme that allow you can convert \r\n" +
    "Multiple Excel files and multiple Folder \r\nto PDF in same Location.";
            // 
            // FrmMain
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(580, 220);
            this.Controls.Add(this.labelDescription);
            this.Controls.Add(this.labelFooter2);
            this.Controls.Add(this.labelFooter1);
            this.Controls.Add(this.IconError);
            this.Controls.Add(this.labelErrorMessage);
            this.Controls.Add(this.BtnLetsGo);
            this.Controls.Add(this.labelInfo);
            this.Controls.Add(this.labelDragFolder);
            this.Controls.Add(this.BtnLoad);
            this.Controls.Add(this.TxtFolderName);
            this.Controls.Add(this.pictureLogo);
            this.Controls.Add(this.pictureDrag);
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
            ((System.ComponentModel.ISupportInitialize)(this.pictureDrag)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureLogo)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label TxtFolderName;
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
    }
}

