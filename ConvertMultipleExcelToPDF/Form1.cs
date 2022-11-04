using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Linq;

namespace ConvertMultipleExcelToPDF
{
    public partial class FrmMain : Form
    {
        // Array of Excel Files found in Folder
        string[] XLSfiles;
        // Array of Excel Files Dragged Directly in form main
        string[] files;

        // Default Text
        string defaultTxtDrag = "Drag your Excel files ...";
        string defaultTxtLoad = "Chose your folder location ...";
        // Store slected path of Folder browser dialog in variable
        string selected_path;
        // Create fileCount to counting number of Excel files found
        int fileCount = 0;
        // Create addTip variable to store All empty excel files
        string addTip;

        bool excelDragged = false;
        bool ischecked_DragFiles = false;
        bool ischecked_WorkBook = false;

        Excel.Application excelApplication = null;
        Excel.Workbook excelWorkBook = null;
        object paramMissing = Type.Missing;

        public FrmMain() => InitializeComponent();

        // Handle Event Click of Buttton Let's Go
        private void BtnLetsGo_Click(object sender, EventArgs e)
        {
            LabelEmptyXLS.Visible = false;
            addTip = "";
            picDone.Visible = false;
            labelInfo.Text = "...";

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
                if ( XLSfiles == null || string.IsNullOrEmpty(TxtBoxLoad.Text) )
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
                ChangeToDefault(true);
            }
            catch (Exception ex)
            {
                labelErrorMessage.Text = ex.Message.ToString();
                ChangeToDefault(false);
                CloseWorkBook();
                QuitExcel();
            }
        }

        private void ChangeToDefault(bool Done)
        {
            Cursor = Cursors.Default;
            if (ischecked_DragFiles)
                TxtDraggedFiles.Text = defaultTxtDrag;
            else
                TxtBoxLoad.Text = defaultTxtLoad;
            if (Done)
            {
                labelInfo.Text = "Done";
                picDone.Visible = true;
            }
            else
            {
                LabelEmptyXLS.Visible = false;
                picDone.Visible = false;
                labelInfo.Text = "...";
                IconError.Visible = false;
            }
        }

        // Handle Event Click of Buttton Load Folder
        private void BtnLoad_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FD = new FolderBrowserDialog();
            if (selected_path != null)
                FD.SelectedPath = selected_path;
            if (FD.ShowDialog() == DialogResult.OK)
            {
                HideControls("Load");
                string path = FD.SelectedPath;
                selected_path = path;
                TxtBoxLoad.Text = path;
                fileCount = SearchXLSFiles(path, out XLSfiles);
                labelInfo.Text = fileCount + " Excel files found";
            }
        }

        private void HideControls(string DragOrLoad)
        {
            labelErrorMessage.Text = string.Empty;
            IconError.Visible = false;
            picDone.Visible = false;
            LabelEmptyXLS.Visible = false;
            if(DragOrLoad == "Drag")
            {
                TxtDraggedFiles.Visible = false;
                TxtBoxLoad.Visible = false;
            }
        }

        // Activate Drag & Drop in Form Main ...
        private void FrmMain_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
            ShowImgDrag(true);
            HideControls("Drag");
        }

        private void FrmMain_DragDrop(object sender, DragEventArgs e)
        {
            ShowImgDrag(false);

            if (ischecked_DragFiles)
            {
                TxtDraggedFiles.Visible = true;
                // Handle event Drag the Excel files.
                files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    string extensionfile = Path.GetExtension(file).ToLower();

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
                    TxtDraggedFiles.Text = "Excel Files was Dragged correctly.";   
                    labelInfo.Text = files.Length + " Excel files found";
                }

                else
                {
                    TxtDraggedFiles.Text = "No Excel Files was Dragged";
                    labelInfo.Text = "...";
                    files = null;
                }
            }

            else
            {
                TxtBoxLoad.Visible = true;
                // Handle event Drag the Folder.
                string path = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
                if (Directory.Exists(path))
                {
                    TxtBoxLoad.Text = path;
                    fileCount = SearchXLSFiles(path, out XLSfiles);
                    labelInfo.Text = fileCount + " Excel files found";
                }
                else
                {
                    TxtBoxLoad.Text = "No Folder was Dragged";
                    labelInfo.Text = "...";
                    XLSfiles = null;
                }
            }
        }

        private void FrmMain_DragLeave(object sender, EventArgs e)
        {
            ShowImgDrag(false);
            labelInfo.Text = "...";

            if (ischecked_DragFiles)
            {
                TxtDraggedFiles.Visible = true;
                TxtDraggedFiles.Text = defaultTxtDrag;
            }
            else
            {
                TxtBoxLoad.Visible = true;
                TxtBoxLoad.Text = defaultTxtLoad;
            }
            XLSfiles = null;
            files = null;
        }

        private void ShowImgDrag(bool condition)
        {
            pictureDrag.Visible = condition ? true : false;
            LoadingImage.Visible = condition ? true : false;
            labelInfo.Visible = condition ? false : true;
        }

        // Handle Methode Search in all Sub-Directory and Get all Excel files found,
        // and bring out to the string array
        private int SearchXLSFiles(string path, out string[] XLSfiles)
        {
            XLSfiles = Directory
                        .GetFiles(path, "*.*", SearchOption.AllDirectories)
                        .Where(s => s.ToLower().EndsWith(".xls") || s.ToLower().EndsWith(".xlsx"))
                        .ToArray();
            return XLSfiles.Length;
        }

        // Start Methode ProcessFiles
        private void ProcessFiles(string[] XLSfiles, string[] files)
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
        private void CloseWorkBook()
        {
            if (excelWorkBook != null)
            {
                // Close the workbook object.
                excelWorkBook.Close(false, paramMissing, paramMissing);
                excelWorkBook = null;
            }
        }

        // Start Method QuitExcel
        private void QuitExcel()
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
            HideControls("CheckBox");
            labelInfo.Text = "...";
            XLSfiles = null;
            files = null;
            addTip = "";

            TxtBoxLoad.Visible = checkBoxDragFiles.Checked ? false : true;
            TxtDraggedFiles.Visible = checkBoxDragFiles.Checked ? true : false;
            ischecked_DragFiles = checkBoxDragFiles.Checked ? true : false;
            BtnLoad.Enabled = checkBoxDragFiles.Checked ? false : true;
            labelDragFolder.Font = checkBoxDragFiles.Checked ? 
                                    new Font(labelDragFolder.Font, FontStyle.Strikeout) :
                                    new Font(labelDragFolder.Font, FontStyle.Regular);
            if(checkBoxDragFiles.Checked)
                TxtDraggedFiles.Text = defaultTxtDrag;
            else
                TxtBoxLoad.Text = defaultTxtLoad;
        }

        // Start Method StartConvert
        private void StartConvert(string[] ExcelFiles)
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
                        else
                        {
                            LabelEmptyXLS.Visible = true;
                            addTip += fname + "\r\n";
                            toolTipDrag.SetToolTip(LabelEmptyXLS, addTip);
                        }
                    }
                    else
                    {
                        // Fix Exception HRESULT : 0x800A03EC
                        // Check the Empty Active worksheet
                        if (dataCount == 0)
                        {
                            LabelEmptyXLS.Visible = true;
                            addTip += fname + "\r\n";
                            toolTipDrag.SetToolTip(LabelEmptyXLS, addTip);
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

        private bool IsEmptyWorkbook(Excel.Workbook wb)
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

        private void LabelEmptyXLS_Click(object sender, EventArgs e)
        {
            MessageBox.Show(addTip,"Empty Excel files",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }
    }
}
