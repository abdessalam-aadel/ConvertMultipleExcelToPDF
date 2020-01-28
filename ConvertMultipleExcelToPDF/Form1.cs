using System;
using System.Windows.Forms;
using System.IO;
using MultipleExcelToPDF;
using System.Diagnostics;

namespace ConvertMultipleExcelToPDF
{
    public partial class FrmMain : Form
    {
        string[] XLSfiles;
        int fileCount = 0;

        public FrmMain()
        {
            InitializeComponent();
        }
        // Handle Event Click of Buttton Let's Go
        private void BtnLetsGo_Click(object sender, EventArgs e)
        {
            if (XLSfiles == null || string.IsNullOrEmpty(TxtFolderName.Text))
            {
                IconError.Visible = true;
                labelErrorMessage.Text = "No source folder has been selected. Please select one.";
                return;
            }
            else if (XLSfiles.Length == 0)
            {
                IconError.Visible = true;
                labelErrorMessage.Text = "No Excel files have been found in the selected folder";
                return;
            }
            Cursor = Cursors.WaitCursor;
            labelInfo.Text = "Processing ...";

            try
            {
                LibFunctions.ProcessFiles(XLSfiles);
                labelInfo.Text = "Done!";
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                labelErrorMessage.Text = ex.Message.ToString();
                LibFunctions.CloseWorkBook();
                LibFunctions.QuitExcel();
            }
        }
        // Handle Event Click of Buttton Load Folder
        private void BtnLoad_Click(object sender, EventArgs e)
        {
            labelErrorMessage.Text = string.Empty;
            pictureDrag.Visible = false;
            IconError.Visible = false;
            FolderBrowserDialog FD = new FolderBrowserDialog();
            if (FD.ShowDialog() == DialogResult.OK)
            {
                TxtFolderName.Text = FD.SelectedPath;
                fileCount = LibFunctions.SearchDirectoryTree(FD.SelectedPath, out XLSfiles);
                labelInfo.Text = fileCount + " Excel files found";
            }
        }
        // Activate Drag & Drop in Form Main ...
        private void FrmMain_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
            pictureDrag.Visible = true;
            labelErrorMessage.Text = string.Empty;
            IconError.Visible = false;
        }
        private void FrmMain_DragDrop(object sender, DragEventArgs e)
        {
            pictureDrag.Visible = false;
            string path = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
            if (Directory.Exists(path))
            {
                labelErrorMessage.Text = string.Empty;
                IconError.Visible = false;
                TxtFolderName.Text = path;
                fileCount = LibFunctions.SearchDirectoryTree(path, out XLSfiles);
                labelInfo.Text = fileCount + " Excel files found";
            }
        }
        private void FrmMain_DragLeave(object sender, EventArgs e)
        {
            pictureDrag.Visible = false;
        }

        private void LinkGit_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://github.com/abdessalam-aadel/ConvertMultipleExcelToPDF");
        }
    }
}
