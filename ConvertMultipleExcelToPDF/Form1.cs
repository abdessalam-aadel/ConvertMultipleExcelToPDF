using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Text.RegularExpressions;

namespace ConvertMultipleExcelToPDF
{
    public partial class FrmMain : Form
    {
        // Array of Excel Files found in Folder
        public static string[] XLSfiles;
        // Array of Excel Files Dragged Directly in form main
        public static string[] files;
        string selected_path;

        int fileCount = 0;
        public static bool excelDragged = false;
        public static bool ischecked_WorkBook = false;
        public static bool ischecked_DragFiles = false;
        private static Excel.Application excelApplication = null;
        private static Excel.Workbook excelWorkBook = null;
        private static object paramMissing = Type.Missing;

        public FrmMain() => InitializeComponent();

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
                if (ischecked_DragFiles)
                    TxtFolderName.Text = "Drag your Excel files ...";
                else
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
            if (selected_path != null)
                FD.SelectedPath = selected_path;
            if (FD.ShowDialog() == DialogResult.OK)
            {
                string path = FD.SelectedPath;
                selected_path = path;
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
            labelInfo.Text = string.Empty;
            IconError.Visible = false;
            TxtFolderName.Visible = false;
            LoadingImage.Visible = true;
        }

        private void FrmMain_DragDrop(object sender, DragEventArgs e)
        {
            pictureDrag.Visible = false;
            LoadingImage.Visible = false;
            TxtFolderName.Visible = true;

            if (ischecked_DragFiles)
            {
                // Handle event Drag the Excel files.
                files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    string extensionfile = Path.GetExtension(file);

                    if (extensionfile == ".xls" || extensionfile == ".xlsx")
                        excelDragged = true;
                    else
                    {
                        excelDragged = false;
                        break;
                    }
                }

                if (excelDragged)
                {
                    TxtFolderName.Text = "Excel Files was Dragged correctly.";
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
                // Handle event Drag the Folder.
                string path = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
                if (Directory.Exists(path))
                {
                    labelErrorMessage.Text = string.Empty;
                    IconError.Visible = false;
                    TxtFolderName.Text = path;
                    fileCount = SearchDirectoryTree(path, out XLSfiles);
                    labelInfo.Text = fileCount + " Excel files found";
                }
                else
                {
                    TxtFolderName.Text = "No Folder was Dragged";
                    labelInfo.Text = "...";
                    XLSfiles = null;
                }
            }
        }

        private void FrmMain_DragLeave(object sender, EventArgs e)
        {
            pictureDrag.Visible = false;
            LoadingImage.Visible = false;
            TxtFolderName.Visible = true;
            labelInfo.Text = "...";
        }

        // Handle Methode Search Directory and Get all Excel files found,
        // and bring out to the string array
        public static int SearchDirectoryTree(string path, out string[] XLSfiles)
        {
            XLSfiles = Directory.GetFiles(path, "*.xls", SearchOption.AllDirectories);
            return XLSfiles.Length;
        }

        // Start Methode ProcessFiles
        public static void ProcessFiles(string[] XLSfiles, string[] files)
        {
            // Condition to separat : Folder or Excel files was Dragged
            if (ischecked_DragFiles)
                StartConvert(files);
            else
                StartConvert(XLSfiles);

            // Exit Excel
            QuitExcel();
        }

        // Start Method CloseWorkBook
        public static void CloseWorkBook()
        {
            if (excelWorkBook != null)
            {
                // Close the workbook object.
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

            // Force garbage collection.
            GC.Collect();
            // Wait for all finalizers to complete before continuing.
            // Without this call to GC.WaitForPendingFinalizers,
            // the worker loop below might execute at the same time
            // as the finalizers. 
            // With this call, the worker loop executes only after
            // all finalizers have been called.
            GC.WaitForPendingFinalizers();
            // Clear string array
            XLSfiles = null;
            files = null;
        }

        private void checkBoxAllWorkBook_CheckedChanged(object sender, EventArgs e)
        {
            ischecked_WorkBook = checkBoxAllWorkBook.Checked ? true : false;
        }

        private void checkBoxDragFiles_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxDragFiles.Checked)
            {
                TxtFolderName.Text = "Drag your Excel files ...";
                labelInfo.Text = "...";
                ischecked_DragFiles = true;
                BtnLoad.Enabled = false;
                labelDragFolder.Font = new Font(labelDragFolder.Font, FontStyle.Strikeout);
                XLSfiles = null;
                files = null;
            }
            else
            {
                TxtFolderName.Text = "Chose your Folder Location ...";
                labelInfo.Text = "...";
                ischecked_DragFiles = false;
                BtnLoad.Enabled = true;
                labelDragFolder.Font = new Font(labelDragFolder.Font, FontStyle.Regular);
                XLSfiles = null;
                files = null;
            }
        }

        // Start Method StartConvert
        public static void StartConvert(string[] ExcelFiles)
        {
            // Creat new instance of Microsoft Excel Application
            excelApplication = new Excel.Application();
            //excelApplication.Visible = false;

            // Declare Parameters :
            // ...
            // XlFixedFormatType object : Specifie whether to save the workbook (PDF format).
            Excel.XlFixedFormatType paramExportFormat = Excel.XlFixedFormatType.xlTypePDF;
            // XlFixedFormatQuality object : Specifie the quality of the exported file (Standard Quality).
            Excel.XlFixedFormatQuality paramExportQuality = Excel.XlFixedFormatQuality.xlQualityStandard;
            // Not open the pdf file after exporting the workbook
            bool paramOpenAfterPublish = false;
            // Include document properties in the exported file
            bool paramIncludeDocProps = true;
            // Ignore any print areas set when exporting
            bool paramIgnorePrintAreas = true;
            // from Object: is the number of the page at which to start exporting,
            // If this parameter is omitted, exporting starts at the beginning.
            object paramFromPage = Type.Missing;
            // to Object: is the number of the last page to export,
            // If this parameter is omitted, exporting ends with the last page.
            object paramToPage = Type.Missing;

            foreach (string filesPath in ExcelFiles)
            {
                FileInfo finfo = new FileInfo(filesPath);
                string fname = finfo.Name;
                // use regular expression to search the filename begin with ~$
                string pattern = @"^~\$";
                // Find matches
                Match m = Regex.Match(fname, pattern);
                if (m.Success)
                    continue;

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

                // Active sheet
                Excel.Worksheet worksheet = excelWorkBook.ActiveSheet;
                // Detect data in all cells of Active sheet
                int dataCount = (int)excelApplication.WorksheetFunction.CountA(worksheet.Cells);

                if (excelWorkBook != null)
                {
                    if (ischecked_WorkBook)
                    {
                        // Check Entire workbook was not Empty
                        if (!IsEmptyWorkbook(excelWorkBook))
                            excelWorkBook.ExportAsFixedFormat(paramExportFormat,
                            paramExportFilePath, paramExportQuality,
                            paramIncludeDocProps, paramIgnorePrintAreas, paramFromPage,
                            paramToPage, paramOpenAfterPublish,
                            paramMissing); // Convert Entire WorkBook to PDF
                    }
                    else
                    {
                        // Fix Exception HRESULT : 0x800A03EC
                        // Check the Empty Active worksheet
                        if (dataCount == 0)
                        {
                            // All cells on the Active worksheet are empty.
                            // -- Skip over --
                            // continues with the next iteration of the loop for-each
                            continue;
                        }

                        // There is at least one cell on the worksheet that has non-empty contents.
                        else
                        {
                            excelWorkBook.ActiveSheet.ExportAsFixedFormat(paramExportFormat,
                            paramExportFilePath, paramExportQuality,
                            paramIncludeDocProps, paramIgnorePrintAreas, paramFromPage,
                            paramToPage, paramOpenAfterPublish,
                            paramMissing); // Convert Active Sheet to PDF
                        }
                    }
                }
                CloseWorkBook();
            }
        }

        public static bool IsEmptyWorkbook(Excel.Workbook wb)
        {
            try
            {
                foreach (Excel.Worksheet sheet in wb.Worksheets)
                {
                    if (excelApplication.WorksheetFunction.CountA(sheet.Cells) != 0)
                        return false;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void ImgGit_Click(object sender, EventArgs e)
        {
            // Go to Github Repository
            Process.Start("https://github.com/abdessalam-aadel/ConvertMultipleExcelToPDF");
        }
    }
}
