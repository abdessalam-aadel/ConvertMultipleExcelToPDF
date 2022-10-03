using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace ConvertMultipleExcelToPDF
{
    public partial class FrmMain : Form
    {
        string[] XLSfiles;
        string[] files;
        int fileCount = 0;
        public static bool excelDragged = false;
        public static bool ischecked_WorkBook = false;
        public static bool ischecked_DragFiles = false;
        private static Excel.Application excelApplication = null;
        private static Excel.Workbook excelWorkBook = null;
        private static object paramMissing = Type.Missing;

        public FrmMain()
        {
            InitializeComponent();
        }

        // Handle Event Click of Buttton Let's Go
        private void BtnLetsGo_Click(object sender, EventArgs e)
        {
            if ( ischecked_DragFiles )
            {
                if ( files == null || !excelDragged )
                {
                    IconError.Visible = true;
                    labelErrorMessage.Text = "No Excel file was Dragged, Try again.";
                    return;
                }
            }

            else
            {
                if ( XLSfiles == null || string.IsNullOrEmpty(TxtFolderName.Text) )
                {
                    IconError.Visible = true;
                    labelErrorMessage.Text = "No source folder was selected, Please select one.";
                    return;
                }

                else if ( XLSfiles.Length == 0 )
                {
                    IconError.Visible = true;
                    labelErrorMessage.Text = "No Excel file was found in the selected folder";
                    return;
                }
            }

            IconError.Visible = false;
            labelErrorMessage.Text = "";
            Cursor = Cursors.WaitCursor;
            labelInfo.Text = "Processing ...";

            try
            {
                ProcessFiles(XLSfiles, files);
                labelInfo.Text = "Done!";
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                labelErrorMessage.Text = ex.Message.ToString();
                Cursor = Cursors.Default;
                labelInfo.Text = "...";
                TxtFolderName.Text = "Chose your Folder Location ...";
                IconError.Visible = false;
                CloseWorkBook();
                QuitExcel();
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
                string path = FD.SelectedPath;
                TxtFolderName.Text = path;
                fileCount = SearchDirectoryTree(path, out XLSfiles);
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

            if (ischecked_DragFiles)
            {
                files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    string extensionfile = Path.GetExtension(file);
                    if (extensionfile != ".xls" && extensionfile != ".xlsx")
                        excelDragged = false;
                    if (extensionfile == ".xls" || extensionfile == ".xlsx")
                        excelDragged = true;
                }

                if (excelDragged)
                {
                    TxtFolderName.Text = "Excel Files was Dragged";
                    labelErrorMessage.Text = string.Empty;
                    IconError.Visible = false;
                    labelInfo.Text = files.Length + " Excel files found";
                }

                else
                {
                    TxtFolderName.Text = "No Excel Files was Dragged";
                    labelInfo.Text = "...";
                }
            }

            else
            {
                string path = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
                if (Directory.Exists(path))
                {
                    labelErrorMessage.Text = string.Empty;
                    IconError.Visible = false;
                    TxtFolderName.Text = path;
                    fileCount = SearchDirectoryTree(path, out XLSfiles);
                    labelInfo.Text = fileCount + " Excel files found";
                }
            }
        }

        private void FrmMain_DragLeave(object sender, EventArgs e)
        {
            pictureDrag.Visible = false;
        }

        private void LinkGit_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Go to Github Repository
            Process.Start("https://github.com/abdessalam-aadel/ConvertMultipleExcelToPDF");
        }

        public static int SearchDirectoryTree(string path, out string[] XLSfiles)
        {
            XLSfiles = Directory.GetFiles(path, "*.xls", SearchOption.AllDirectories);
            return XLSfiles.Length;
        }

        public static void ProcessFiles(string[] XLSfiles, string[] files)
        {
            excelApplication = new Excel.Application();
            Excel.XlFixedFormatType paramExportFormat = Excel.XlFixedFormatType.xlTypePDF;
            Excel.XlFixedFormatQuality paramExportQuality = Excel.XlFixedFormatQuality.xlQualityStandard;
            bool paramOpenAfterPublish = false; // not open the pdf file after publish 
            bool paramIncludeDocProps = true;
            bool paramIgnorePrintAreas = true;
            object paramFromPage = Type.Missing;
            object paramToPage = Type.Missing;

            if (ischecked_DragFiles)
            {
                foreach (string filesPath in files)
                {
                    string paramSourceBookPath = filesPath;
                    // Get Extension of filePath
                    string extension = Path.GetExtension(filesPath);
                    string paramExportFilePath = filesPath.Replace(extension, ".pdf"); // Replace Extension .xls or .xlsx

                    // Open the source workbook.
                    excelWorkBook = excelApplication.Workbooks.Open(paramSourceBookPath,
                    paramMissing, paramMissing, paramMissing, paramMissing,
                            paramMissing, paramMissing, paramMissing, paramMissing,
                            paramMissing, paramMissing, paramMissing, paramMissing,
                            paramMissing, paramMissing);

                    // Save it in the target format.
                    if (excelWorkBook != null)
                    {
                        if (ischecked_WorkBook)
                            excelWorkBook.ExportAsFixedFormat(paramExportFormat,
                            paramExportFilePath, paramExportQuality,
                            paramIncludeDocProps, paramIgnorePrintAreas, paramFromPage,
                            paramToPage, paramOpenAfterPublish,
                            paramMissing); // Convert Entire WorkBook to PDF
                        else
                            excelWorkBook.ActiveSheet.ExportAsFixedFormat(paramExportFormat,
                            paramExportFilePath, paramExportQuality,
                            paramIncludeDocProps, paramIgnorePrintAreas, paramFromPage,
                            paramToPage, paramOpenAfterPublish,
                            paramMissing); // Convert Active Sheet(s) to PDF
                    }
                    CloseWorkBook();
                }
            }

            else
            {
                foreach (string filePath in XLSfiles)
                {
                    string paramSourceBookPath = filePath;
                    // Get Extension of filePath
                    string extension = Path.GetExtension(filePath);
                    string paramExportFilePath = filePath.Replace(extension, ".pdf"); // Replace Extension .xls or .xlsx

                    // Open the source workbook.
                    excelWorkBook = excelApplication.Workbooks.Open(paramSourceBookPath,
                    paramMissing, paramMissing, paramMissing, paramMissing,
                            paramMissing, paramMissing, paramMissing, paramMissing,
                            paramMissing, paramMissing, paramMissing, paramMissing,
                            paramMissing, paramMissing);

                    // Save it in the target format.
                    if (excelWorkBook != null)
                    {
                        if (ischecked_WorkBook)
                            excelWorkBook.ExportAsFixedFormat(paramExportFormat,
                            paramExportFilePath, paramExportQuality,
                            paramIncludeDocProps, paramIgnorePrintAreas, paramFromPage,
                            paramToPage, paramOpenAfterPublish,
                            paramMissing); // Convert Entire WorkBook to PDF
                        else
                            excelWorkBook.ActiveSheet.ExportAsFixedFormat(paramExportFormat,
                            paramExportFilePath, paramExportQuality,
                            paramIncludeDocProps, paramIgnorePrintAreas, paramFromPage,
                            paramToPage, paramOpenAfterPublish,
                            paramMissing); // Convert Active Sheet(s) to PDF
                    }
                    CloseWorkBook();
                }
            }
            
            QuitExcel();
        }

        // Start Method CloseWorkBook
        public static void CloseWorkBook()
        {
            // Close the workbook object.
            if (excelWorkBook != null)
            {
                excelWorkBook.Close(false, paramMissing, paramMissing);
                excelWorkBook = null;
            }
        }

        // Start Method QuitExcel
        public static void QuitExcel()
        {
            // Quit Excel and release the ApplicationClass object.
            if (excelApplication != null)
            {
                excelApplication.Quit();
                excelApplication = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void checkBoxAllWorkBook_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAllWorkBook.Checked)
                ischecked_WorkBook = true;
            else
                ischecked_WorkBook = false;
        }

        private void checkBoxDragFiles_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxDragFiles.Checked)
            {
                ischecked_DragFiles = true;
                BtnLoad.Enabled = false;
                labelDragFolder.Font = new Font(labelDragFolder.Font, FontStyle.Strikeout);
            }
            else
            {
                ischecked_DragFiles = false;
                BtnLoad.Enabled = true;
                labelDragFolder.Font = new Font(labelDragFolder.Font, FontStyle.Regular);
            }
        }
    }
}
