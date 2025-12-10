using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.WindowsAPICodePack.Dialogs;
using MS.WindowsAPICodePack.Internal;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Point = System.Drawing.Point;

namespace LChart_Comparison_Tool
{
    public partial class Form1 : Form
    {
        List<ExcelRangeBase> navigateLeft = new List<ExcelRangeBase>();
        List<ExcelRangeBase> navigateRight = new List<ExcelRangeBase>();
        List<ExcelRangeBase> navigateUp = new List<ExcelRangeBase>();
        List<ExcelRangeBase> navigateDown = new List<ExcelRangeBase>();
        const string SheetLChartNovel = "L-Chart(NOVEL)";
        const string SheetLChart = "L-Chart";
        const string SheetManual = "Manual Sheet";

        public List<ExcelRangeBase> ParentMergedCells = new();

        // Internal queues → BFS recursion
        private Queue<(int row, int col)> UpQueue = new();
        private Queue<(int row, int col)> DownQueue = new();


        int newRowToWriteAt = 7;
        private string outputFolder = "";

        public Form1()
        {
            InitializeComponent();
            comboBox1.DrawMode = DrawMode.Normal;
            comboBox1.ForeColor = Color.Black;
            comboBox1.BackColor = Color.White;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox1.Items.Add("All");
            comboBox1.Items.Add("Operation No. Master");
            comboBox1.Items.Add("Operation No and EM Task");
            comboBox1.Items.Add("EM Task Master");
            comboBox1.Items.Add("EM Task and Equipment");
            comboBox1.Items.Add("Equipment Master");
            comboBox1.Items.Add("EM Task and Tool Group");
            comboBox1.Items.Add("Tool Grpoup Master");
            comboBox1.Items.Add("EM Task and Part");
            comboBox1.Items.Add("Part Master");
            comboBox1.Items.Add("Operation No and MR");
            comboBox1.Items.Add("Parent and Child Master");
            comboBox1.Items.Add("Tool Master");
            comboBox1.SelectedIndex = 0;

        }
        int x = 0;
        public List<inputInfo> inputData = new List<inputInfo>();

        private void Form1_Resize(object sender, EventArgs e)
        {
            // Trigger repaint when form is resized
            this.Invalidate();
        }
        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            // Create a pen with color and thickness
            Pen blackPen = new Pen(Color.Green, 3);

            // Define start and end points
            Point startPoint = new Point(50, 100);
            Point endPoint = new Point(300, 100);


            int x = label3.Left + label3.Width / 2;
            int yStart = label3.Bottom + 5;
            int yEnd = linkLabel3.Top + linkLabel3.Height / 2;


            // Draw the line
            e.Graphics.DrawLine(blackPen, x, yStart, x, yEnd);


            int y = linkLabel1.Top + linkLabel1.Height / 2;
            int xStart = x;
            int xEnd = linkLabel1.Left - 3;

            e.Graphics.DrawLine(blackPen, xStart, y, xEnd, y);

            y = linkLabel2.Top + linkLabel2.Height / 2;
            xStart = x;
            xEnd = linkLabel2.Left - 3;

            e.Graphics.DrawLine(blackPen, xStart, y, xEnd, y);

            y = linkLabel3.Top + linkLabel3.Height / 2;
            xStart = x;
            xEnd = linkLabel3.Left;

            e.Graphics.DrawLine(blackPen, xStart, y, xEnd, y);

            //e.Graphics.DrawLine(blackPen, x, yStart,x,yEnd);
        }

        private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
        {
            System.Windows.Forms.TabControl tc = sender as System.Windows.Forms.TabControl;
            this.BackColor = Color.FromArgb(13, 91, 155);
            if (tc == null) return;
            // Custom colors per tab index,
            Color[] tabColors = { Color.FromArgb(13, 91, 155), Color.FromArgb(21, 104, 65), Color.FromArgb(190, 80, 20), Color.FromArgb(128, 120, 15), Color.FromArgb(236, 103, 32), Color.Teal };
            using (Brush backBrush = new SolidBrush(tabColors[e.Index % tabColors.Length]))
            using (Brush textBrush = new SolidBrush(Color.FloralWhite))
            {
                e.Graphics.FillRectangle(backBrush, e.Bounds);
                e.Graphics.DrawString(tc.TabPages[e.Index].Text, e.Font, textBrush, e.Bounds.X + 5, e.Bounds.Y + 5);
            }
        }
        private void Helpbtn_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                Application pptApp = new Application();
                pptApp.Visible = MsoTriState.msoTrue;

                Presentations presentations = pptApp.Presentations;
                string xlOutputPath = System.IO.Directory.GetCurrentDirectory() + @"\bin\SOP_Gate1_Checklist_Automation.pptx";
                Presentation presentation = presentations.Open(xlOutputPath,
                    MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
                pptApp.Activate();
                this.Cursor = Cursors.Default;
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        private void BrowseBtn_Click(object sender, EventArgs e)
        {
            if (label12.Text == "LChart Portal Folder :")
            {

                CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                dialog.IsFolderPicker = true;
                dialog.InitialDirectory = "C:\\";

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    string selectedPath = dialog.FileName;
                    //MessageBox.Show("Selected Folder: " + selectedPath);
                    textBox1.Text = selectedPath;
                }
            }
            else if (label12.Text == "LChart Portal File :")
            {
                using (OpenFileDialog fileDlg = new OpenFileDialog())
                {
                    fileDlg.Title = "Select Master File";
                    fileDlg.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm|All Files|*.*";

                    if (fileDlg.ShowDialog() == DialogResult.OK)
                    {
                        textBox1.Text = fileDlg.FileName;
                    }
                }
            }
        }

        private void FolderBrowse_Click(object sender, EventArgs e)
        {

            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            dialog.InitialDirectory = "C:\\";

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string selectedPath = dialog.FileName;
                //MessageBox.Show("Selected Folder: " + selectedPath);
                InputPath.Text = selectedPath;
            }
        }

        public class TableInfoNew
        {
            public int StrtRow = 0;
            public MinMax m_oATA = null;
            public MinMax m_oDesc = null;
            public MinMax m_oReplace = null;
            public MinMax m_oRpReason = null;
            public MinMax m_oComment = null;
            public TableInfoNew()
            {
                m_oATA = new MinMax();
                m_oDesc = new MinMax();
                m_oReplace = new MinMax();
                m_oRpReason = new MinMax();
                m_oComment = new MinMax();
            }
            public void copyMinMax()
            {
                m_oDesc.copyMinMax();
                m_oATA.copyMinMax();
                m_oReplace.copyMinMax();
                m_oRpReason.copyMinMax();
                m_oComment.copyMinMax();
            }
        }
        public class MinMax
        {
            public double m_dMin = 0.0;
            public double m_dMax = 0.0;
            public double m_dMinNew = 0.0;
            public double m_dMaxNew = 0.0;
            public void copyMinMax()
            {
                m_dMinNew = m_dMin;
                m_dMaxNew = m_dMax;
            }
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                if (comboBox1.SelectedItem.ToString() == "All")
                {
                    label12.Text = "LChart Portal Folder :";
                }
                else
                {
                    label12.Text = "LChart Portal File :";
                }
            }
        }
        private void label12_Click(object sender, EventArgs e)
        {

        }
        private void InputPath_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string path = textBox1.Text.Trim();
            string selectedName = comboBox1.Text.Trim();

            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(selectedName))
                return;


            if (selectedName == "All")
            {

            }
            else if (path.IndexOf(selectedName, StringComparison.OrdinalIgnoreCase) >= 0)
            {
            }
            else
            {
                MessageBox.Show(
                            $"The selected master '{selectedName}' does not match the given LChart portal file path.",
                            "Path Mismatch",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );
            }
        }

        private void Form1_Checklist_Click(object sender, EventArgs e)
        {

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar1.Visible = false;
            label16.Visible = false;
            //cmbDirection.SelectedIndex = 0;
        }
        private void Form1_Update_Click(object sender, EventArgs e)
        {
            try
            {
                string portalPath = textBox1.Text.Trim(); string inputFolder = InputPath.Text.Trim(); string excelName = comboBox1.Text.Trim();

                if (string.IsNullOrWhiteSpace(inputFolder) && string.IsNullOrWhiteSpace(portalPath))
                {
                    MessageBox.Show("Please select an input folder and LChart portal file paths.",
                                    "Missing Path",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    textBox1.Focus();
                    return;
                }
                if (string.IsNullOrWhiteSpace(inputFolder))
                {
                    MessageBox.Show("Please select an input folder path.",
                                    "Missing Path",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    InputPath.Focus();
                    return;
                }
                if (!System.IO.Directory.Exists(inputFolder))
                {
                    MessageBox.Show("The selected input folder path does not exist.",
                                    "Invalid Path",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    InputPath.Focus();
                    return;
                }


                if (string.IsNullOrWhiteSpace(portalPath))
                {
                    MessageBox.Show("Please select a LChart portal path.",
                                    "Missing Path",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    textBox1.Focus();
                    return;
                }

                if (!System.IO.File.Exists(portalPath) && !System.IO.Directory.Exists(portalPath))
                {
                    MessageBox.Show("The selected portal file or folder path does not exist.",
                                    "Invalid Path",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    textBox1.Focus();
                    return;
                }
                this.Cursor = Cursors.WaitCursor;
                progressBar1.Visible = true;
                label16.Visible = true;
                progressBar1.Value = 0;
                x = 2;
                Progressed(x);
                if (InputPath.Text.Length > 0)
                {
                    string[] files = Directory.GetFiles(inputFolder, "*.xlsx");
                    if (files.Length == 0)
                    {
                        MessageBox.Show("No Excel files found in the selected input folder.");
                        return;
                    }
                    //UpdateMasterWithInputFiles(portalPath, inputFolder, excelName);
                    if (excelName == "All")
                    {
                        MessageBox.Show("All master files updated successfully.", "LChart master creation tool");
                    }
                    else
                    {
                        MessageBox.Show(excelName + " file updated successfully.", "LChart master creation tool");


                    }
                    this.Cursor = Cursors.Default;
                }
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                label16.Visible = false;
                progressBar1.Visible = false;
                progressBar1.Value = 0;
                label16.Text = "";
                label10.Text = "";
                MessageBox.Show("Error: " + ex.Message, "Excel Creator",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                label16.Visible = false;
                progressBar1.Visible = false;
                progressBar1.Value = 0;
                label16.Text = "";
                this.Cursor = Cursors.Default;
            }
        }
        private async void Progressed(int x)
        {
            progressBar1.Value += x;
            progressBar1.Refresh();
            System.Windows.Forms.Application.DoEvents();
            await Task.Delay(50);
        }


        //private void UpdateMasterWithInputFiles(string portalPath, string inputFolder, string excelName)
        //{
        //    string portalFolder = string.Empty;
        //    string portalFilePath = string.Empty;
        //    string[] outputFiles;

        //    // --- Detect if portalPath is a folder or a single Excel file ---
        //    if (Directory.Exists(portalPath))
        //    {
        //        // Case 1: User selected a folder → take all Excel files in it
        //        portalFolder = portalPath;
        //        outputFiles = Directory.GetFiles(portalPath, "*.xlsx");
        //    }
        //    else if (File.Exists(portalPath))
        //    {
        //        // Case 2: User selected a single file → take only that one
        //        portalFolder = Path.GetDirectoryName(portalPath);
        //        portalFilePath = portalPath;
        //        outputFiles = new string[] { portalFilePath };
        //    }
        //    else
        //    {
        //        MessageBox.Show("Invalid portal path. Please select a valid folder or Excel file.");
        //        return;
        //    }

        //    // --- Validate input folder ---
        //    if (!Directory.Exists(inputFolder))
        //    {
        //        MessageBox.Show("Input folder not found.");
        //        return;
        //    }
        //    string[] inputFiles = Directory.GetFiles(inputFolder, "*.xlsx");
        //    if (inputFiles.Length == 0)
        //    {
        //        MessageBox.Show("No input Excel files found in the selected folder.");
        //        return;
        //    }
        //    string outputFolder = checkSavingFolder(excelName);
        //    string masterFilePath = Path.Combine(outputFolder, excelName + ".xlsx");

        //    if (!excelName.Equals("All", StringComparison.OrdinalIgnoreCase))
        //    {
        //        progressBar1.Maximum = 6;
        //        var filesInFolder = Directory.GetFiles(outputFolder, "*.xlsx");
        //        foreach (var file in filesInFolder)
        //        {
        //            string fileName = Path.GetFileNameWithoutExtension(file);
        //            if (!fileName.Equals(excelName, StringComparison.OrdinalIgnoreCase))
        //            {
        //                try { File.Delete(file); } catch { }
        //            }
        //            else
        //            {
        //                using (var package = new ExcelPackage(new FileInfo(file)))
        //                {
        //                    var ws = package.Workbook.Worksheets.FirstOrDefault();
        //                    if (ws != null && ws.Dimension != null)
        //                    {
        //                        int lastRow = ws.Dimension.End.Row;
        //                        if (lastRow >= 7)
        //                            ws.DeleteRow(7, lastRow - 6);
        //                    }
        //                    package.Save();
        //                }
        //            }
        //        }

        //        if (!File.Exists(masterFilePath))
        //        {
        //            using (var package = new ExcelPackage(new FileInfo(masterFilePath)))
        //            {
        //                string sheetName = getJapaneseSheetName(excelName);
        //                var ws = package.Workbook.Worksheets.Add(sheetName);
        //                FillHeaderRows(ws, excelName);
        //                package.Save();
        //            }
        //        }

        //        // Step 3: Read input files → Fill data list
        //        inputData.Clear();
        //        x = 2;
        //        Progressed(x);
        //        label16.Text = "Reading Cyient LChart Module files....";
        //        readInputFiles(inputFolder, inputFiles, excelName);
        //        x = 2;
        //        Progressed(x);
        //        label16.Text = "Printing data to selected master file....";
        //        this.Refresh();
        //        Thread.Sleep(5000);
        //        // Step 4: Write data into master
        //        WriteInputDataToMaster(portalFilePath, masterFilePath, getSheetName(excelName));
        //    }
        //    else
        //    {
        //        progressBar1.Maximum = 28;
        //        x = 2;
        //        Progressed(x);
        //        label16.Text = "Reading Cyient LChart Module files....";
        //        inputData.Clear();
        //        readInputFiles(inputFolder, inputFiles, excelName);
        //        // Step 3 (All masters)
        //        string[] allMasterNames = {
        //        "Operation No and MR", "Operation No. Master", "Operation No and EM Task",
        //        "EM Task Master", "EM Task and Tool Group", "EM Task and Part",
        //        "EM Task and Equipment", "Equipment Master", "Tool Master",
        //        "Part Master", "Parent and Child Master", "Tool Grpoup Master"};

        //        string outputDetails = string.Empty;
        //        foreach (var currentExcelName in allMasterNames)
        //        {
        //            x = 2;
        //            Progressed(x);
        //            label16.Text = "Creating L-Chart master for  " + currentExcelName + "....";
        //            string currentMasterPath = Path.Combine(outputFolder, currentExcelName + ".xlsx");
        //            string currentPortalPath = Directory.GetFiles(portalFolder, "*.xlsx")
        //                .FirstOrDefault(f =>
        //                {
        //                    string fileName = Path.GetFileNameWithoutExtension(f);
        //                    // Match base name ignoring trailing timestamp numbers
        //                    return fileName.StartsWith(currentExcelName, StringComparison.OrdinalIgnoreCase);
        //                });

        //            CheckandcreateOutput(outputFolder, currentExcelName, currentMasterPath);

        //            WriteInputDataToMaster(currentPortalPath, currentMasterPath, getSheetName(currentExcelName));
        //        }


        //    }
        //}


        //public void CheckandcreateOutput(string outputFolder, string currentExcelName, string currentMasterPath)
        //{
        //    string outputdetails = string.Empty;

        //    if (!File.Exists(currentMasterPath))
        //    {
        //        using (var package = new ExcelPackage(new FileInfo(currentMasterPath)))
        //        {
        //            string sheetName = getJapaneseSheetName(currentExcelName);
        //            var ws = package.Workbook.Worksheets.Add(sheetName);
        //            FillHeaderRows(ws, currentExcelName);
        //            package.Save();
        //        }
        //    }
        //    else
        //    {
        //        using (var package = new ExcelPackage(new FileInfo(currentMasterPath)))
        //        {
        //            var ws = package.Workbook.Worksheets.FirstOrDefault();
        //            if (ws != null && ws.Dimension != null)
        //            {
        //                int lastRow = ws.Dimension.End.Row;
        //                if (lastRow >= 7)
        //                    ws.DeleteRow(7, lastRow - 6);
        //            }
        //            package.Save();
        //        }
        //    }


        //}

        //private void WriteInputDataToMaster(string portalFilePath, string masterFilePath, string excelName)
        //{
        //    using (var package = new ExcelPackage(new FileInfo(masterFilePath)))
        //    {
        //        var ws = package.Workbook.Worksheets.FirstOrDefault();
        //        if (ws == null) return;

        //        ExcelRangeBase dim = ws.Cells[ws.Dimension.Address];
        //        if (dim == null) return;

        //        int startRow = 7;
        //        int headerRow = 3;
        //        int lastCol = dim.End.Column;

        //        // ---------------- Step A: Build column map ----------------
        //        var columnMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        //        for (int col = 1; col <= lastCol; col++)
        //        {
        //            string header = ws.Cells[headerRow, col].Text.Trim();
        //            if (!string.IsNullOrEmpty(header) && !columnMap.ContainsKey(header))
        //                columnMap[header] = col;
        //        }

        //        // ---------------- Step B: Header to Property Mapping ----------------
        //        var headerToProperty = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        //{
        //    {"Block No", "Block_No"}, {"Eng Type", "Eng_Type"}, {"Module", "Module"},
        //    {"OFF or ON", "OFF_OR_ON"}, {"Operation Name", "Operation_Name"},
        //    {"Operation No", "Operation_No"}, {"Edit Type", "Edit_Type"}, {"MR", "MR"},
        //    {"EM Task", "EM_Task"}, {"EM Task Step", "EM_Task_Step"}, {"Tool Group", "Tool_Group"},
        //    {"Equipment Utility", "Equipment_Utility"}, {"QTY", "QTY"}, {"ATA No", "ATA_No"},
        //    {"Fig. Item No", "Fig_Item_No"}, {"P/N", "PN"}, {"Description", "Description"},
        //    {"Parent Operation No", "PARENT_NO"}, {"Revision Note", "Note"}
        //};

        //        var type = typeof(inputInfo);

        //        // ---------------- Step C: Write uniqueData to master ----------------
        //        var uniqueData = inputData.Where(x => x.SheetName == excelName).ToList();
        //        if (uniqueData.Count > 0)
        //        {
        //            int writeRow = startRow;
        //            foreach (var item in uniqueData)
        //            {
        //                foreach (var header in columnMap.Keys)
        //                {
        //                    if (headerToProperty.TryGetValue(header, out string propName))
        //                    {
        //                        var prop = type.GetProperty(propName);
        //                        if (prop != null)
        //                            ws.Cells[writeRow, columnMap[header]].Value = prop.GetValue(item)?.ToString();
        //                    }
        //                }
        //                writeRow++;
        //            }
        //        }

        //        // ---------------- Step D: Read portal file ----------------
        //        List<Dictionary<string, string>> portalData = new List<Dictionary<string, string>>();
        //        using (var portalPkg = new ExcelPackage(new FileInfo(portalFilePath)))
        //        {
        //            var portalWs = portalPkg.Workbook.Worksheets.FirstOrDefault();
        //            if (portalWs != null && portalWs.Dimension != null)
        //            {
        //                int endRow = portalWs.Dimension.End.Row;
        //                int endCol = portalWs.Dimension.End.Column;
        //                var validHeaders = columnMap.Keys.Where(h => headerToProperty.ContainsKey(h)).ToList();

        //                object[,] data = portalWs.Cells[startRow, 1, endRow, endCol].Value as object[,];
        //                if (data != null)
        //                {
        //                    int rows = data.GetLength(0);
        //                    int cols = data.GetLength(1);
        //                    for (int r = 0; r < rows; r++)
        //                    {
        //                        var rowDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //                        foreach (string header in validHeaders)
        //                        {
        //                            int idx = columnMap[header] - 1;
        //                            if (idx < cols)
        //                                rowDict[header] = data[r, idx]?.ToString()?.Trim() ?? "";
        //                        }
        //                        portalData.Add(rowDict);
        //                    }
        //                }
        //            }
        //        }

        //        // ---------------- Step E: Compare and Update ----------------
        //        int masterLastRow = ws.Dimension.End.Row;
        //        int editTypeCol = columnMap.ContainsKey("Edit Type") ? columnMap["Edit Type"] : -1;
        //        int noteCol = columnMap.ContainsKey("Note") ? columnMap["Note"] : -1;
        //        int revisionNoteCol = columnMap.ContainsKey("Revision Note") ? columnMap["Revision Note"] : -1;

        //        var compareProps = type.GetProperties()
        //            .Where(p => p.Name != "SheetName")
        //            .Select(p => p.Name).ToList();

        //        // Cache master data
        //        var masterData = new List<Dictionary<string, string>>(masterLastRow - startRow + 1);
        //        object[,] masterValues = ws.Cells[startRow, 1, masterLastRow, lastCol].Value as object[,];
        //        if (masterValues != null)
        //        {
        //            int rows = masterValues.GetLength(0);
        //            int cols = masterValues.GetLength(1);
        //            for (int r = 0; r < rows; r++)
        //            {
        //                var rowDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //                foreach (var header in headerToProperty.Keys)
        //                {
        //                    if (columnMap.TryGetValue(header, out int colIdx) && colIdx <= cols)
        //                        rowDict[header] = masterValues[r, colIdx - 1]?.ToString()?.Trim() ?? "";
        //                }
        //                masterData.Add(rowDict);
        //            }
        //        }

        //        // ---------------- Enhancement: Added Note check ----------------
        //        List<int> deleteAddedRows = new List<int>();

        //        for (int i = 0; i < masterData.Count; i++)
        //        {
        //            if (noteCol > 0 &&
        //                masterData[i].TryGetValue("Note", out string noteValue) &&
        //                noteValue.Equals("added", StringComparison.OrdinalIgnoreCase))
        //            {
        //                bool isMatched = portalData.Any(pd =>
        //                    headerToProperty.Keys.All(h =>
        //                        (!masterData[i].ContainsKey(h) && !pd.ContainsKey(h)) ||
        //                        (masterData[i].ContainsKey(h) && pd.ContainsKey(h) &&
        //                         string.Equals(masterData[i][h], pd[h], StringComparison.OrdinalIgnoreCase))));

        //                if (isMatched)
        //                    deleteAddedRows.Add(i + startRow);
        //                else if (revisionNoteCol > 0)
        //                    ws.Cells[i + startRow, revisionNoteCol].Value = "ADDED (New Part)";
        //            }
        //        }

        //        // Delete matched "added" rows
        //        deleteAddedRows.Sort();
        //        deleteAddedRows.Reverse();
        //        foreach (int r in deleteAddedRows)
        //            ws.DeleteRow(r);
        //        // ---------------- End Enhancement ----------------

        //        // Continue with rest of Step E logic (existing compare)
        //        HashSet<string> masterKeys = new HashSet<string>(
        //            masterData.Select(md => string.Join("|", compareProps.Select(p =>
        //            {
        //                var header = headerToProperty.FirstOrDefault(x => x.Value == p).Key;
        //                return header != null && md.ContainsKey(header) ? md[header] : "";
        //            }))),
        //            StringComparer.OrdinalIgnoreCase);

        //        HashSet<string> portalKeys = new HashSet<string>(
        //            portalData.Select(pd => string.Join("|", compareProps.Select(p =>
        //            {
        //                var header = headerToProperty.FirstOrDefault(x => x.Value == p).Key;
        //                return header != null && pd.ContainsKey(header) ? pd[header] : "";
        //            }))),
        //            StringComparer.OrdinalIgnoreCase);

        //        // Mark "C" for changed rows
        //        var rowsToMark = new List<int>();
        //        for (int i = 0; i < masterData.Count; i++)
        //        {
        //            var row = masterData[i];
        //            string key = string.Join("|", compareProps.Select(p =>
        //            {
        //                var header = headerToProperty.FirstOrDefault(x => x.Value == p).Key;
        //                return header != null && row.ContainsKey(header) ? row[header] : "";
        //            }));

        //            if (!portalKeys.Contains(key))
        //                rowsToMark.Add(i + startRow);
        //        }

        //        foreach (int r in rowsToMark)
        //        {
        //            ws.Cells[r, editTypeCol].Value = "C";
        //            ws.Cells[r, editTypeCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            ws.Cells[r, editTypeCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //        }

        //        // Remove identical rows
        //        List<int> rowsToDelete = new List<int>();
        //        for (int i = 0; i < masterData.Count; i++)
        //        {
        //            string mKey = string.Join("|", compareProps.Select(p =>
        //            {
        //                var header = headerToProperty.FirstOrDefault(x => x.Value == p).Key;
        //                return header != null && masterData[i].ContainsKey(header) ? masterData[i][header] : "";
        //            }));
        //            if (portalKeys.Contains(mKey))
        //                rowsToDelete.Add(i + startRow);
        //        }

        //        rowsToDelete.Sort();
        //        rowsToDelete.Reverse();
        //        foreach (int r in rowsToDelete)
        //            ws.DeleteRow(r);
        //        // ✅ Clear non-“ADDED (New Part)” Revision Notes after all checks
        //        if (revisionNoteCol > 0)
        //        {
        //            int totalRows = ws.Dimension.End.Row;
        //            for (int r = startRow; r <= totalRows; r++)
        //            {
        //                var val = ws.Cells[r, revisionNoteCol].Text?.Trim();
        //                if (!string.Equals(val, "ADDED (New Part)", StringComparison.OrdinalIgnoreCase))
        //                    ws.Cells[r, revisionNoteCol].Clear();
        //            }
        //        }

        //        // ----------------Skip comparison if file name contains "and" ----------------
        //        string fileNameOnly = Path.GetFileNameWithoutExtension(masterFilePath);
        //        if (fileNameOnly.IndexOf("and", StringComparison.OrdinalIgnoreCase) <= 0)
        //        {
        //            ws.Cells[ws.Dimension.Address].AutoFitColumns();

        //            // ✅ Clear non-“ADDED (New Part)” Revision Notes
        //            if (revisionNoteCol > 0)
        //            {
        //                int totalRows = ws.Dimension.End.Row;
        //                for (int r = startRow; r <= totalRows; r++)
        //                {
        //                    var val = ws.Cells[r, revisionNoteCol].Text?.Trim();
        //                    if (!string.Equals(val, "ADDED (New Part)", StringComparison.OrdinalIgnoreCase))
        //                        ws.Cells[r, revisionNoteCol].Clear();
        //                }
        //            }

        //            package.Save();
        //            return;
        //        }

        //        // ---------------- Add missing rows (D) ----------------
        //        int newRow = masterLastRow + 1;
        //        var addedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        //        foreach (var portalRow in portalData)
        //        {
        //            string pKey = string.Join("|", compareProps.Select(p =>
        //            {
        //                var header = headerToProperty.FirstOrDefault(x => x.Value == p).Key;
        //                return header != null && portalRow.ContainsKey(header) ? portalRow[header] : "";
        //            }));

        //            if (!masterKeys.Contains(pKey) && !addedKeys.Contains(pKey))
        //            {
        //                foreach (var header in columnMap.Keys)
        //                {
        //                    if (portalRow.TryGetValue(header, out string val))
        //                        ws.Cells[newRow, columnMap[header]].Value = val;
        //                }
        //                if (editTypeCol > 0)
        //                {
        //                    ws.Cells[newRow, editTypeCol].Value = "D";
        //                    ws.Cells[newRow, editTypeCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //                    ws.Cells[newRow, editTypeCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //                }
        //                newRow++;
        //                addedKeys.Add(pKey);
        //            }
        //        }
        //        if (revisionNoteCol > 0)
        //        {
        //            int totalRows = ws.Dimension.End.Row;
        //            for (int r = startRow; r <= totalRows; r++)
        //            {
        //                var val = ws.Cells[r, revisionNoteCol].Text?.Trim();
        //                if (!string.IsNullOrEmpty(val))
        //                    ws.Cells[r, revisionNoteCol].Clear();
        //            }
        //        }


        //        ws.Cells[ws.Dimension.Address].AutoFitColumns();
        //        package.Save();
        //    }
        //}


        //private void readInputFiles(string inputFolder, string[] files, string excelName)
        //{

        //    string sheetName = string.Empty;
        //    List<string> allSheetNames = getSheetNames();
        //    foreach (string filePath in files)
        //    {
        //        using (var package = new ExcelPackage(new FileInfo(filePath)))
        //        {
        //            if (excelName == "All")
        //            {

        //                for (int i = 0; i < allSheetNames.Count; i++)
        //                {
        //                    sheetName = allSheetNames[i].Trim();
        //                    createInputList(package, sheetName);
        //                }

        //            }
        //            else
        //            {
        //                sheetName = getSheetName(excelName);
        //                createInputList(package, sheetName);

        //            }

        //        }

        //    }
        //}



        //private void createInputList(ExcelPackage package, string sheetName)
        //{

        //    string name = string.Empty;
        //    if (sheetName == "All")
        //    {
        //        foreach (var worksheet in package.Workbook.Worksheets)
        //        {
        //            name = worksheet.Name;
        //            readExcelSheet(package, name);
        //        }

        //    }
        //    else if (sheetName != "All")
        //    {
        //        name = sheetName;
        //        readExcelSheet(package, name);
        //    }
        //}



        //private void readExcelSheet(ExcelPackage package, string name)
        //{
        //    var uniqueSet = new HashSet<string>();
        //    var worksheet = package.Workbook.Worksheets[name];
        //    if (worksheet != null)
        //    {
        //        int lastRow = worksheet.Cells
        //        .Where(c => !string.IsNullOrEmpty(c.Text))
        //        .Select(c => c.Start.Row)
        //        .DefaultIfEmpty(0)
        //        .Max();

        //        for (int row = 7; row <= lastRow; row++)
        //        {
        //            inputInfo input = new inputInfo();
        //            var c = worksheet.Cells;
        //            input.SheetName = name;
        //            input.Block_No = c[row, 1].Text;
        //            input.Eng_Type = c[row, 2].Text;
        //            input.Module = c[row, 3].Text;


        //            switch (name)
        //            {
        //                case "MR Sheet":
        //                    input.OFF_OR_ON = c[row, 4].Text;
        //                    input.Operation_Name = c[row, 5].Text;
        //                    input.Edit_Type = c[row, 6].Text;
        //                    input.Operation_No = c[row, 7].Text;
        //                    input.MR = c[row, 8].Text;
        //                    input.Note = c[row, 10].Text;
        //                    break;
        //                case "Manual Sheet":
        //                    input.OFF_OR_ON = c[row, 4].Text;
        //                    input.Operation_Name = c[row, 5].Text;
        //                    input.Edit_Type = c[row, 6].Text;
        //                    input.Operation_No = c[row, 7].Text;
        //                    input.EM_Task = c[row, 8].Text;
        //                    input.EM_Task_Step = c[row, 9].Text;
        //                    input.Note = c[row, 11].Text;
        //                    break;
        //                case "Equipment Utility Sheet":
        //                    input.OFF_OR_ON = c[row, 4].Text;
        //                    input.Operation_No = c[row, 5].Text;
        //                    input.Operation_Name = c[row, 6].Text;
        //                    input.Edit_Type = c[row, 7].Text;
        //                    input.EM_Task = c[row, 8].Text;
        //                    input.EM_Task_Step = c[row, 9].Text;
        //                    input.Equipment_Utility = c[row, 10].Text;
        //                    input.Note = c[row, 12].Text;
        //                    break;
        //                case "Tools Sheet":
        //                    input.OFF_OR_ON = c[row, 4].Text;
        //                    input.Operation_No = c[row, 5].Text;
        //                    input.Operation_Name = c[row, 6].Text;
        //                    input.Edit_Type = c[row, 7].Text;
        //                    input.EM_Task = c[row, 8].Text;
        //                    input.EM_Task_Step = c[row, 9].Text;
        //                    input.Tool_Group = c[row, 10].Text;
        //                    input.QTY = c[row, 11].Text;
        //                    input.Note = c[row, 13].Text;
        //                    break;
        //                case "Parts Sheet":
        //                    input.OFF_OR_ON = c[row, 4].Text;
        //                    input.Operation_No = c[row, 5].Text;
        //                    input.Operation_Name = c[row, 6].Text;
        //                    input.Edit_Type = c[row, 7].Text;
        //                    input.EM_Task = c[row, 8].Text;
        //                    input.EM_Task_Step = c[row, 9].Text;
        //                    input.ATA_No = c[row, 10].Text;
        //                    input.Fig_Item_No = c[row, 11].Text;
        //                    input.PN = c[row, 12].Text;
        //                    input.Description = c[row, 13].Text;
        //                    input.QTY = c[row, 14].Text;
        //                    input.Note = c[row, 16].Text;
        //                    break;

        //                case "Link Sheet":
        //                    input.Block_No = c[row, 1].Text;
        //                    input.Operation_No = c[row, 2].Text;
        //                    input.Operation_Name = c[row, 3].Text;
        //                    input.PARENT_NO = c[row, 5].Text;
        //                    input.Note = c[row, 7].Text;
        //                    break;
        //            }
        //            string uniqueKey = $"{input.SheetName}|{input.Block_No}|{input.Eng_Type}|{input.Module}|" +
        //                $"{input.OFF_OR_ON}|{input.Operation_No}|{input.Operation_Name}|";


        //            if (!uniqueSet.Contains(uniqueKey))
        //            {
        //                uniqueSet.Add(uniqueKey);
        //                inputData.Add(input);
        //            }
        //        }
        //    }

        //}


        private List<string> getSheetNames()
        {

            List<string> allExcelNames = new List<string>
            {
                "MR Sheet",
                "Manual Sheet",
                "Equipment Utility Sheet",
                "Tools Sheet",
                "Parts Sheet",
                "Link Sheet",

            };
            return allExcelNames;
        }
        private string getSheetName(string excelName)
        {

            switch (excelName.Trim())
            {
                case "Operation No and MR":
                    return "MR Sheet";
                case "Operation No. Master":
                    return "Manual Sheet";
                case "Operation No and EM Task":
                    return "Manual Sheet";
                case "EM Task Master":
                    return "Manual Sheet";
                case "EM Task and Equipment":
                    return "Equipment Utility Sheet";
                case "Equipment Master":
                    return "Equipment Utility Sheet";
                case "Tool Grpoup Master":
                    return "Tools Sheet";
                case "EM Task and Tool Group":
                    return "Tools Sheet";
                case "Tool Master":
                    return "Tools Sheet";
                case "Part Master":
                    return "Parts Sheet";
                case "EM Task and Part":
                    return "Parts Sheet";
                case "Parent and Child Master":
                    return "Link Sheet";

                default:
                    return "";
            }
        }

        public class inputInfo
        {
            public string SheetName { get; set; }
            public string Block_No { get; set; }
            public string Eng_Type { get; set; }
            public string Module { get; set; }
            public string OFF_OR_ON { get; set; }
            public string Operation_Name { get; set; }
            public string Edit_Type { get; set; }
            public string Operation_No { get; set; }
            public string MR { get; set; }
            public string EM_Task { get; set; }
            public string EM_Task_Step { get; set; }
            public string Equipment_Utility { get; set; }
            public string Tool_Group { get; set; }
            public string QTY { get; set; }
            public string Tool { get; set; }
            public string Tool_Name { get; set; }
            public string ATA_No { get; set; }
            public string Fig_Item_No { get; set; }
            public string PN { get; set; }
            public string Description { get; set; }
            public string PARENT_NO { get; set; }
            public string Note { get; set; }
        }
        private static string checkSavingFolder(string excelName)
        {

            string debugPath = System.IO.Directory.GetCurrentDirectory();
            string toolName = "LChartMasters";
            string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");

            string[] existingFolders = Directory.GetDirectories(debugPath, toolName + "*");

            string outputFolder;

            if (existingFolders.Length == 0)
            {
                outputFolder = Path.Combine(debugPath, toolName + "_" + timeStamp);
                Directory.CreateDirectory(outputFolder);
            }
            else
            {
                string latestFolder = existingFolders.OrderByDescending(f => Directory.GetCreationTime(f)).First();
                string newFolderName = Path.Combine(debugPath, toolName + "_" + timeStamp);

                try
                {
                    Directory.Move(latestFolder, newFolderName);
                }
                catch (Exception)
                {
                    // Fallback if rename fails — create new instead
                    newFolderName = Path.Combine(debugPath, toolName + "_" + timeStamp + "_new");
                    Directory.CreateDirectory(newFolderName);
                }

                outputFolder = newFolderName;
            }

            if (string.IsNullOrEmpty(excelName))
                excelName = "Output";  // fallback

            return outputFolder;
        }
        private string getJapaneseSheetName(string excelName)
        {
            switch (excelName)
            {
                case "Operation No. Master":
                    return "作業Noマスタテーブル";
                case "Operation No and MR":
                    return "作業No-MR 紐づけテーブル";
                case "Operation No and EM Task":
                    return "作業No-EM Task 紐づけテーブル";
                case "EM Task Master":
                    return "EM Taskマスタテーブル";
                case "EM Task and Equipment":
                    return "EM Task-設備 紐づけテーブル";
                case "Equipment Master":
                    return "設備マスタテーブル";
                case "Tool Grpoup Master":
                    return "治工具グループマスタテーブル";
                case "Part Master":
                    return "部品マスタテーブル";
                case "Parent and Child Master":
                    return "親子関係テーブル";
                case "Tool Master":
                    return "治工具マスタテーブル";
                case "EM Task and Part":
                    return "EM Task-部品 紐づけテーブル";
                case "EM Task and Tool Group":
                    return "EM Task-治工具 紐づけテーブル";
                default:
                    return " ";
            }
        }


        //private void FillHeaderRows(ExcelWorksheet ws, string excelName)
        //{

        //    switch (excelName)
        //    {
        //        case "Operation No. Master":
        //            ws.Cells[1, 13].Value = "画像フォルダパス：";
        //            ws.Cells[1, 14].Value = "XXX";
        //            ws.Cells[2, 1].Value = "Edit Type";
        //            ws.Cells[2, 2].Value = "作業No";
        //            ws.Cells[2, 3].Value = "作業名称";
        //            ws.Cells[2, 4].Value = "ブロックNo";
        //            ws.Cells[2, 5].Value = "工程";
        //            ws.Cells[2, 6].Value = "ワーク姿勢";
        //            ws.Cells[2, 7].Value = "占有面積";
        //            ws.Cells[2, 8].Value = "作業人数";
        //            ws.Cells[2, 9].Value = "C/T";
        //            ws.Cells[2, 10].Value = "機種";
        //            ws.Cells[2, 11].Value = "Module";
        //            ws.Cells[2, 12].Value = "OFF or ON";
        //            ws.Cells[2, 13].Value = "属性値(セレクタ)";
        //            ws.Cells[2, 14].Value = "画像";
        //            ws.Cells[2, 15].Value = "Note";
        //            ws.Cells[2, 16].Value = "Revision Number";
        //            ws.Cells[2, 17].Value = "改訂理由";
        //            ws.Cells[2, 18].Value = "登録日";
        //            ws.Cells[2, 19].Value = "登録者";
        //            ws.Cells[2, 20].Value = "無効日";
        //            ws.Cells[3, 1].Value = "Edit Type";
        //            ws.Cells[3, 2].Value = "Operation No";
        //            ws.Cells[3, 3].Value = "Operation Name";
        //            ws.Cells[3, 4].Value = "Block No";
        //            ws.Cells[3, 5].Value = "Process";
        //            ws.Cells[3, 6].Value = "Position";
        //            ws.Cells[3, 7].Value = "Area";
        //            ws.Cells[3, 8].Value = "Manpower";
        //            ws.Cells[3, 9].Value = "C/T";
        //            ws.Cells[3, 10].Value = "Eng Type";
        //            ws.Cells[3, 11].Value = "Module";
        //            ws.Cells[3, 12].Value = "OFF or ON";
        //            ws.Cells[3, 13].Value = "Selector";
        //            ws.Cells[3, 14].Value = "Image path";
        //            ws.Cells[3, 15].Value = "Note";
        //            ws.Cells[3, 16].Value = "Revision Number";
        //            ws.Cells[3, 17].Value = "Revision Note";
        //            ws.Cells[3, 18].Value = "Revised Date";
        //            ws.Cells[3, 19].Value = "Revised by";
        //            ws.Cells[3, 20].Value = "Invalid Date";
        //            ws.Cells[4, 1].Value = "C";
        //            ws.Cells[5, 1].Value = "U";
        //            ws.Cells[6, 1].Value = "D";
        //            ws.Cells[4, 2].Value = "*";
        //            ws.Cells[5, 2].Value = "*";
        //            ws.Cells[6, 2].Value = "*";
        //            ws.Cells[4, 3].Value = "*";
        //            ws.Cells[4, 13].Value = "*";
        //            ws.Cells[4, 14].Value = "*";
        //            ws.Cells[4, 15].Value = "*";
        //            ws.Cells[4, 16].Value = "*";
        //            ws.Cells[4, 17].Value = "*";
        //            ws.Cells[4, 16].Value = "-";
        //            ws.Cells[5, 16].Value = "-";
        //            ws.Cells[6, 16].Value = "-";
        //            ws.Cells[4, 18].Value = "-";
        //            ws.Cells[5, 18].Value = "-";
        //            ws.Cells[6, 18].Value = "-";
        //            ws.Cells[4, 19].Value = "-";
        //            ws.Cells[5, 19].Value = "-";
        //            ws.Cells[6, 19].Value = "-";
        //            ws.Cells[4, 20].Value = "-";
        //            ws.Cells[5, 20].Value = "-";
        //            ws.Cells[6, 20].Value = "-";

        //            // ==== Formatting (EPPlus Style) ====
        //            var allCells = ws.Cells["A1:T6"];
        //            allCells.Style.Font.Name = "Calibri";
        //            allCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //            allCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //            // Font sizes
        //            ws.Row(1).Style.Font.Size = 11;
        //            ws.Row(2).Style.Font.Size = 11;
        //            ws.Row(3).Style.Font.Size = 10;
        //            ws.Row(4).Style.Font.Size = 11;
        //            ws.Row(5).Style.Font.Size = 11;
        //            ws.Row(6).Style.Font.Size = 11;

        //            ws.Cells[2, 1, 6, 20].AutoFitColumns();

        //            // Background colors
        //            ws.Cells["A2:T3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A2:T3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 198, 220));

        //            ws.Cells["A4:T6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A4:T6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(244, 245, 250));
        //            break;

        //        case "Operation No and MR":
        //            ws.Cells[1, 1].Value = "Reference情報";
        //            ws.Cells[2, 1].Value = "ブロックNo";
        //            ws.Cells[2, 2].Value = "機種";
        //            ws.Cells[2, 3].Value = "Module名";
        //            ws.Cells[2, 4].Value = "OFF or ON";
        //            ws.Cells[2, 5].Value = "作業名称";
        //            ws.Cells[2, 6].Value = "Edit Type";
        //            ws.Cells[2, 7].Value = "作業No";
        //            ws.Cells[2, 8].Value = "MR";
        //            ws.Cells[2, 9].Value = "Note";
        //            ws.Cells[2, 10].Value = "改訂理由";
        //            ws.Cells[2, 11].Value = "登録日";
        //            ws.Cells[2, 12].Value = "登録者";
        //            ws.Cells[2, 13].Value = "更新日";
        //            ws.Cells[2, 14].Value = "更新者";
        //            ws.Cells[2, 15].Value = "無効日";
        //            ws.Cells[2, 16].Value = "無効者";
        //            ws.Cells[3, 1].Value = "Block No";
        //            ws.Cells[3, 2].Value = "Eng Type";
        //            ws.Cells[3, 3].Value = "Module";
        //            ws.Cells[3, 4].Value = "OFF or ON";
        //            ws.Cells[3, 5].Value = "Operation Name";
        //            ws.Cells[3, 6].Value = "Edit Type";
        //            ws.Cells[3, 7].Value = "Operation No";
        //            ws.Cells[3, 8].Value = "MR";
        //            ws.Cells[3, 9].Value = "Note";
        //            ws.Cells[3, 10].Value = "Revision Note";
        //            ws.Cells[3, 11].Value = "Revised Date";
        //            ws.Cells[3, 12].Value = "Revised by";
        //            ws.Cells[3, 13].Value = "Updated Date";
        //            ws.Cells[3, 14].Value = "Updated by";
        //            ws.Cells[3, 15].Value = "Invalid Date";
        //            ws.Cells[3, 16].Value = "Invalid by";
        //            ws.Cells[4, 6].Value = "C";
        //            ws.Cells[5, 6].Value = "U";
        //            ws.Cells[6, 6].Value = "D";
        //            ws.Cells[4, 7].Value = "*";
        //            ws.Cells[5, 7].Value = "*";
        //            ws.Cells[6, 7].Value = "*";
        //            ws.Cells[4, 8].Value = "*";
        //            ws.Cells[5, 8].Value = "*";
        //            ws.Cells[6, 8].Value = "*";
        //            ws.Cells[4, 11].Value = "-";
        //            ws.Cells[5, 11].Value = "-";
        //            ws.Cells[6, 11].Value = "-";
        //            ws.Cells[4, 12].Value = "-";
        //            ws.Cells[5, 12].Value = "-";
        //            ws.Cells[6, 12].Value = "-";
        //            ws.Cells[4, 13].Value = "-";
        //            ws.Cells[5, 13].Value = "-";
        //            ws.Cells[6, 13].Value = "-";
        //            ws.Cells[4, 14].Value = "-";
        //            ws.Cells[5, 14].Value = "-";
        //            ws.Cells[6, 14].Value = "-";
        //            ws.Cells[4, 15].Value = "-";
        //            ws.Cells[5, 15].Value = "-";
        //            ws.Cells[6, 15].Value = "-";
        //            ws.Cells[4, 16].Value = "-";
        //            ws.Cells[5, 16].Value = "-";
        //            ws.Cells[6, 16].Value = "-";

        //            // ==== Formatting (EPPlus Style) ====
        //            allCells = ws.Cells["A1:P6"];
        //            allCells.Style.Font.Name = "Calibri";
        //            allCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //            allCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //            // Font sizes
        //            ws.Row(1).Style.Font.Size = 20;
        //            ws.Row(2).Style.Font.Size = 11;
        //            ws.Row(3).Style.Font.Size = 10;
        //            ws.Row(4).Style.Font.Size = 11;
        //            ws.Row(5).Style.Font.Size = 11;
        //            ws.Row(6).Style.Font.Size = 11;

        //            ws.Cells[1, 1].Style.Font.Bold = true;
        //            ws.Cells["A1:B1"].Merge = true;
        //            ws.Cells[2, 1, 6, 16].AutoFitColumns();
        //            ws.Column(1).Width = 15;

        //            // Background colors
        //            ws.Cells["A2:E3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A2:E3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 188, 188));

        //            ws.Cells["A1:E1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["A4:E6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A4:E6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["F2:P3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["F2:P3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 198, 220));

        //            ws.Cells["F4:P6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["F4:P6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(244, 245, 250));
        //            break;

        //        case "Operation No and EM Task":
        //            ws.Cells[1, 1].Value = "Reference情報";
        //            ws.Cells[2, 1].Value = "ブロックNo";
        //            ws.Cells[2, 2].Value = "機種";
        //            ws.Cells[2, 3].Value = "Module名";
        //            ws.Cells[2, 4].Value = "OFF or ON";
        //            ws.Cells[2, 5].Value = "作業名称";
        //            ws.Cells[2, 6].Value = "Edit Type";
        //            ws.Cells[2, 7].Value = "作業No";
        //            ws.Cells[2, 8].Value = "EM Task";
        //            ws.Cells[2, 9].Value = "EM Task Step";
        //            ws.Cells[2, 10].Value = "Note";
        //            ws.Cells[2, 11].Value = "改訂理由";
        //            ws.Cells[2, 12].Value = "登録日";
        //            ws.Cells[2, 13].Value = "登録者";
        //            ws.Cells[2, 14].Value = "更新日";
        //            ws.Cells[2, 15].Value = "更新者";
        //            ws.Cells[2, 16].Value = "無効日";
        //            ws.Cells[2, 17].Value = "無効者";
        //            ws.Cells[3, 1].Value = "Block No";
        //            ws.Cells[3, 2].Value = "Eng Type";
        //            ws.Cells[3, 3].Value = "Module";
        //            ws.Cells[3, 4].Value = "OFF or ON";
        //            ws.Cells[3, 5].Value = "Operation Name";
        //            ws.Cells[3, 6].Value = "Edit Type";
        //            ws.Cells[3, 7].Value = "Operation No";
        //            ws.Cells[3, 8].Value = "EM Task";
        //            ws.Cells[3, 9].Value = "EM Task Step";
        //            ws.Cells[3, 10].Value = "Note";
        //            ws.Cells[3, 11].Value = "Revision Note";
        //            ws.Cells[3, 12].Value = "Revised Date";
        //            ws.Cells[3, 13].Value = "Revised by";
        //            ws.Cells[3, 14].Value = "Updated Date";
        //            ws.Cells[3, 15].Value = "Updated by";
        //            ws.Cells[3, 16].Value = "Invalid Date";
        //            ws.Cells[3, 17].Value = "Invalid by";
        //            ws.Cells[4, 6].Value = "C";
        //            ws.Cells[5, 6].Value = "U";
        //            ws.Cells[6, 6].Value = "D";
        //            ws.Cells[4, 7].Value = "*";
        //            ws.Cells[5, 7].Value = "*";
        //            ws.Cells[6, 7].Value = "*";
        //            ws.Cells[4, 8].Value = "*";
        //            ws.Cells[5, 8].Value = "*";
        //            ws.Cells[6, 8].Value = "*";
        //            ws.Cells[4, 9].Value = "*";
        //            ws.Cells[5, 9].Value = "*";
        //            ws.Cells[6, 9].Value = "*";
        //            ws.Cells[4, 12].Value = "-";
        //            ws.Cells[5, 12].Value = "-";
        //            ws.Cells[6, 12].Value = "-";
        //            ws.Cells[4, 13].Value = "-";
        //            ws.Cells[5, 13].Value = "-";
        //            ws.Cells[6, 13].Value = "-";
        //            ws.Cells[4, 14].Value = "-";
        //            ws.Cells[5, 14].Value = "-";
        //            ws.Cells[6, 14].Value = "-";
        //            ws.Cells[4, 15].Value = "-";
        //            ws.Cells[5, 15].Value = "-";
        //            ws.Cells[6, 15].Value = "-";
        //            ws.Cells[4, 16].Value = "-";
        //            ws.Cells[5, 16].Value = "-";
        //            ws.Cells[6, 16].Value = "-";
        //            ws.Cells[4, 17].Value = "-";
        //            ws.Cells[5, 17].Value = "-";
        //            ws.Cells[6, 17].Value = "-";

        //            // ==== Formatting (EPPlus Style) ====
        //            allCells = ws.Cells["A1:Q6"];
        //            allCells.Style.Font.Name = "Calibri";
        //            allCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //            allCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //            // Font sizes
        //            ws.Row(1).Style.Font.Size = 20;
        //            ws.Row(2).Style.Font.Size = 11;
        //            ws.Row(3).Style.Font.Size = 10;
        //            ws.Row(4).Style.Font.Size = 11;
        //            ws.Row(5).Style.Font.Size = 11;
        //            ws.Row(6).Style.Font.Size = 11;

        //            // Bold cell A1
        //            ws.Cells[1, 1].Style.Font.Bold = true;
        //            ws.Cells["A1:B1"].Merge = true;
        //            ws.Cells[2, 1, 6, 17].AutoFitColumns();
        //            ws.Column(1).Width = 15;

        //            // Background colors
        //            ws.Cells["A2:E3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A2:E3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 188, 188));

        //            ws.Cells["A1:E1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["A4:E6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A4:E6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["F2:Q3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["F2:Q3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 198, 220));

        //            ws.Cells["F4:Q6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["F4:Q6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(244, 245, 250));

        //            break;

        //        case "EM Task Master":
        //            ws.Cells[1, 1].Value = "Reference情報";
        //            ws.Cells[2, 1].Value = "ブロックNo";
        //            ws.Cells[2, 2].Value = "機種";
        //            ws.Cells[2, 3].Value = "Module名";
        //            ws.Cells[2, 4].Value = "OFF or ON";
        //            ws.Cells[2, 5].Value = "作業No";
        //            ws.Cells[2, 6].Value = "作業名称";
        //            ws.Cells[2, 7].Value = "Edit Type";
        //            ws.Cells[2, 8].Value = "EM Task";
        //            ws.Cells[2, 9].Value = "EM Task Step"; ;
        //            ws.Cells[2, 10].Value = "Note";
        //            ws.Cells[2, 11].Value = "Revision Number";
        //            ws.Cells[2, 12].Value = "改訂理由";
        //            ws.Cells[2, 13].Value = "登録日";
        //            ws.Cells[2, 14].Value = "登録者";
        //            ws.Cells[2, 15].Value = "無効日";
        //            ws.Cells[3, 1].Value = "Block No";
        //            ws.Cells[3, 2].Value = "Eng Type";
        //            ws.Cells[3, 3].Value = "Module";
        //            ws.Cells[3, 4].Value = "OFF or ON";
        //            ws.Cells[3, 5].Value = "Operation No";
        //            ws.Cells[3, 6].Value = "Operation Name";
        //            ws.Cells[3, 7].Value = "Edit Type";
        //            ws.Cells[3, 8].Value = "EM Task";
        //            ws.Cells[3, 9].Value = "EM Task Step";
        //            ws.Cells[3, 10].Value = "Note";
        //            ws.Cells[3, 11].Value = "Revision Number";
        //            ws.Cells[3, 12].Value = "Revision Note";
        //            ws.Cells[3, 13].Value = "Revised Date";
        //            ws.Cells[3, 14].Value = "Revised by";
        //            ws.Cells[3, 15].Value = "Invalid Date";
        //            ws.Cells[4, 7].Value = "C";
        //            ws.Cells[5, 7].Value = "U";
        //            ws.Cells[6, 7].Value = "D";
        //            ws.Cells[4, 8].Value = "*";
        //            ws.Cells[5, 8].Value = "*";
        //            ws.Cells[6, 8].Value = "*";
        //            ws.Cells[4, 9].Value = "*";
        //            ws.Cells[5, 9].Value = "*";
        //            ws.Cells[6, 9].Value = "*";
        //            ws.Cells[4, 11].Value = "-";
        //            ws.Cells[5, 11].Value = "-";
        //            ws.Cells[6, 11].Value = "-";
        //            ws.Cells[4, 13].Value = "-";
        //            ws.Cells[5, 13].Value = "-";
        //            ws.Cells[6, 13].Value = "-";
        //            ws.Cells[4, 14].Value = "-";
        //            ws.Cells[5, 14].Value = "-";
        //            ws.Cells[6, 14].Value = "-";
        //            ws.Cells[4, 15].Value = "-";
        //            ws.Cells[5, 15].Value = "-";
        //            ws.Cells[6, 15].Value = "-";

        //            // ==== Formatting (EPPlus Style) ====
        //            allCells = ws.Cells["A1:O6"];
        //            allCells.Style.Font.Name = "Calibri";
        //            allCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //            allCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //            // Font sizes
        //            ws.Row(1).Style.Font.Size = 20;
        //            ws.Row(2).Style.Font.Size = 11;
        //            ws.Row(3).Style.Font.Size = 10;
        //            ws.Row(4).Style.Font.Size = 11;
        //            ws.Row(5).Style.Font.Size = 11;
        //            ws.Row(6).Style.Font.Size = 11;

        //            ws.Cells[1, 1].Style.Font.Bold = true;
        //            ws.Cells["A1:B1"].Merge = true;
        //            ws.Cells[2, 1, 6, 15].AutoFitColumns();
        //            ws.Column(1).Width = 15;

        //            // Background colors
        //            ws.Cells["A2:F3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A2:F3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 188, 188));

        //            ws.Cells["A1:F1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["A4:F6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A4:F6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["G2:O3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["G2:O3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 198, 220));

        //            ws.Cells["G4:O6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["G4:O6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(244, 245, 250));
        //            break;

        //        case "EM Task and Equipment":
        //            ws.Cells[1, 1].Value = "Reference情報";
        //            ws.Cells[2, 1].Value = "ブロックNo";
        //            ws.Cells[2, 2].Value = "機種";
        //            ws.Cells[2, 3].Value = "Module名";
        //            ws.Cells[2, 4].Value = "OFF or ON";
        //            ws.Cells[2, 5].Value = "作業No";
        //            ws.Cells[2, 6].Value = "作業名称";
        //            ws.Cells[2, 7].Value = "Edit Type";
        //            ws.Cells[2, 8].Value = "EM Task";
        //            ws.Cells[2, 9].Value = "EM Task Step"; ;
        //            ws.Cells[2, 10].Value = "設備";
        //            ws.Cells[2, 11].Value = "Note";
        //            ws.Cells[2, 12].Value = "改訂理由";
        //            ws.Cells[2, 13].Value = "登録日";
        //            ws.Cells[2, 14].Value = "登録者";
        //            ws.Cells[2, 15].Value = "更新日";
        //            ws.Cells[2, 16].Value = "更新者";
        //            ws.Cells[2, 17].Value = "無効日";
        //            ws.Cells[2, 18].Value = "無効者";
        //            ws.Cells[3, 1].Value = "Block No";
        //            ws.Cells[3, 2].Value = "Eng Type";
        //            ws.Cells[3, 3].Value = "Module";
        //            ws.Cells[3, 4].Value = "OFF or ON";
        //            ws.Cells[3, 5].Value = "Operation No";
        //            ws.Cells[3, 6].Value = "Operation Name";
        //            ws.Cells[3, 7].Value = "Edit Type";
        //            ws.Cells[3, 8].Value = "EM Task";
        //            ws.Cells[3, 9].Value = "EM Task Step";
        //            ws.Cells[3, 10].Value = "Equipment Utility";
        //            ws.Cells[3, 11].Value = "Note";
        //            ws.Cells[3, 12].Value = "Revision Note";
        //            ws.Cells[3, 13].Value = "Revised Date";
        //            ws.Cells[3, 14].Value = "Revised by";
        //            ws.Cells[3, 15].Value = "Updated Date";
        //            ws.Cells[3, 16].Value = "Updated by";
        //            ws.Cells[3, 17].Value = "Invalid Date";
        //            ws.Cells[3, 18].Value = "Invalid by";
        //            ws.Cells[4, 7].Value = "C";
        //            ws.Cells[5, 7].Value = "U";
        //            ws.Cells[6, 7].Value = "D";
        //            ws.Cells[4, 8].Value = "*";
        //            ws.Cells[5, 8].Value = "*";
        //            ws.Cells[6, 8].Value = "*";
        //            ws.Cells[4, 9].Value = "*";
        //            ws.Cells[5, 9].Value = "*";
        //            ws.Cells[6, 9].Value = "*";
        //            ws.Cells[4, 10].Value = "*";
        //            ws.Cells[5, 10].Value = "*";
        //            ws.Cells[6, 10].Value = "*";
        //            ws.Cells[4, 13].Value = "-";
        //            ws.Cells[5, 13].Value = "-";
        //            ws.Cells[6, 13].Value = "-";
        //            ws.Cells[4, 14].Value = "-";
        //            ws.Cells[5, 14].Value = "-";
        //            ws.Cells[6, 14].Value = "-";
        //            ws.Cells[4, 15].Value = "-";
        //            ws.Cells[5, 15].Value = "-";
        //            ws.Cells[6, 15].Value = "-";
        //            ws.Cells[4, 16].Value = "-";
        //            ws.Cells[5, 16].Value = "-";
        //            ws.Cells[6, 16].Value = "-";
        //            ws.Cells[4, 17].Value = "-";
        //            ws.Cells[5, 17].Value = "-";
        //            ws.Cells[6, 17].Value = "-";
        //            ws.Cells[4, 18].Value = "-";
        //            ws.Cells[5, 18].Value = "-";
        //            ws.Cells[6, 18].Value = "-";

        //            // ==== Formatting (EPPlus Style) ====
        //            allCells = ws.Cells["A1:R6"];
        //            allCells.Style.Font.Name = "Calibri";
        //            allCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //            allCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //            // Font sizes
        //            ws.Row(1).Style.Font.Size = 20;
        //            ws.Row(2).Style.Font.Size = 11;
        //            ws.Row(3).Style.Font.Size = 10;
        //            ws.Row(4).Style.Font.Size = 11;
        //            ws.Row(5).Style.Font.Size = 11;
        //            ws.Row(6).Style.Font.Size = 11;

        //            // Bold cell A1
        //            ws.Cells[1, 1].Style.Font.Bold = true;
        //            ws.Cells["A1:B1"].Merge = true;
        //            ws.Cells[2, 1, 6, 18].AutoFitColumns();
        //            ws.Column(1).Width = 15;

        //            // Background colors
        //            ws.Cells["A2:F3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A2:F3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 188, 188));

        //            ws.Cells["A1:F1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["A4:F6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A4:F6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["G2:R3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["G2:R3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 198, 220));

        //            ws.Cells["G4:R6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["G4:R6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(244, 245, 250));

        //            break;

        //        case "Equipment Master":
        //            ws.Cells[1, 1].Value = "Reference情報";
        //            ws.Cells[2, 1].Value = "ブロックNo";
        //            ws.Cells[2, 2].Value = "機種";
        //            ws.Cells[2, 3].Value = "Module名";
        //            ws.Cells[2, 4].Value = "OFF or ON";
        //            ws.Cells[2, 5].Value = "作業No";
        //            ws.Cells[2, 6].Value = "作業名称";
        //            ws.Cells[2, 7].Value = "EM Task";
        //            ws.Cells[2, 8].Value = "EM Task Step";
        //            ws.Cells[2, 9].Value = "Edit Type";
        //            ws.Cells[2, 10].Value = "設備";
        //            ws.Cells[2, 11].Value = "Note";
        //            ws.Cells[2, 12].Value = "Revision Number";
        //            ws.Cells[2, 13].Value = "改訂理由";
        //            ws.Cells[2, 14].Value = "登録日";
        //            ws.Cells[2, 15].Value = "登録者";
        //            ws.Cells[2, 16].Value = "無効日";
        //            ws.Cells[3, 1].Value = "Block No";
        //            ws.Cells[3, 2].Value = "Eng Type";
        //            ws.Cells[3, 3].Value = "Module";
        //            ws.Cells[3, 4].Value = "OFF or ON";
        //            ws.Cells[3, 5].Value = "Operation No";
        //            ws.Cells[3, 6].Value = "Operation Name";
        //            ws.Cells[3, 7].Value = "EM Task";
        //            ws.Cells[3, 8].Value = "EM Task Step";
        //            ws.Cells[3, 9].Value = "Edit Type";
        //            ws.Cells[3, 10].Value = "Equipment Utility";
        //            ws.Cells[3, 11].Value = "Note";
        //            ws.Cells[3, 12].Value = "Revision Number";
        //            ws.Cells[3, 13].Value = "Revision Note";
        //            ws.Cells[3, 14].Value = "Revised Date";
        //            ws.Cells[3, 15].Value = "Revised by";
        //            ws.Cells[3, 16].Value = "Invalid Date";
        //            ws.Cells[4, 9].Value = "C";
        //            ws.Cells[5, 9].Value = "U";
        //            ws.Cells[6, 9].Value = "D";
        //            ws.Cells[4, 10].Value = "*";
        //            ws.Cells[5, 10].Value = "*";
        //            ws.Cells[6, 10].Value = "*";
        //            ws.Cells[4, 12].Value = "-";
        //            ws.Cells[5, 12].Value = "-";
        //            ws.Cells[6, 12].Value = "-";
        //            ws.Cells[4, 14].Value = "-";
        //            ws.Cells[5, 14].Value = "-";
        //            ws.Cells[6, 14].Value = "-";
        //            ws.Cells[4, 15].Value = "-";
        //            ws.Cells[5, 15].Value = "-";
        //            ws.Cells[6, 15].Value = "-";
        //            ws.Cells[4, 16].Value = "-";
        //            ws.Cells[5, 16].Value = "-";
        //            ws.Cells[6, 16].Value = "-";

        //            // ==== Formatting (EPPlus Style) ====
        //            allCells = ws.Cells["A1:P6"];
        //            allCells.Style.Font.Name = "Calibri";
        //            allCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //            allCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //            ws.Row(1).Style.Font.Size = 20;
        //            ws.Row(2).Style.Font.Size = 11;
        //            ws.Row(3).Style.Font.Size = 10;
        //            ws.Row(4).Style.Font.Size = 11;
        //            ws.Row(5).Style.Font.Size = 11;
        //            ws.Row(6).Style.Font.Size = 11;

        //            ws.Cells[1, 1].Style.Font.Bold = true;
        //            ws.Cells["A1:B1"].Merge = true;
        //            ws.Cells[2, 1, 6, 16].AutoFitColumns();
        //            ws.Column(1).Width = 15;

        //            ws.Cells["A2:H3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A2:H3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 188, 188));

        //            ws.Cells["A1:H1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A1:H1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["A4:H6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A4:H6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["I2:P3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["I2:P3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 198, 220));

        //            ws.Cells["I4:P6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["I4:P6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(244, 245, 250));
        //            break;

        //        case "Tool Grpoup Master":
        //            ws.Cells[1, 1].Value = "Reference情報";
        //            ws.Cells[2, 1].Value = "ブロックNo";
        //            ws.Cells[2, 2].Value = "機種";
        //            ws.Cells[2, 3].Value = "Module名";
        //            ws.Cells[2, 4].Value = "OFF or ON";
        //            ws.Cells[2, 5].Value = "作業No";
        //            ws.Cells[2, 6].Value = "作業名称";
        //            ws.Cells[2, 7].Value = "EM Task";
        //            ws.Cells[2, 8].Value = "EM Task Step";
        //            ws.Cells[2, 9].Value = "Edit Type";
        //            ws.Cells[2, 10].Value = "治工具グループ";
        //            ws.Cells[2, 11].Value = "治工具";
        //            ws.Cells[2, 12].Value = "Note";
        //            ws.Cells[2, 13].Value = "Revision Number";
        //            ws.Cells[2, 14].Value = "改訂理由";
        //            ws.Cells[2, 15].Value = "登録日";
        //            ws.Cells[2, 16].Value = "登録者";
        //            ws.Cells[2, 17].Value = "無効日";
        //            ws.Cells[3, 1].Value = "Block No";
        //            ws.Cells[3, 2].Value = "Eng Type";
        //            ws.Cells[3, 3].Value = "Module";
        //            ws.Cells[3, 4].Value = "OFF or ON";
        //            ws.Cells[3, 5].Value = "Operation No";
        //            ws.Cells[3, 6].Value = "Operation Name";
        //            ws.Cells[3, 7].Value = "EM Task";
        //            ws.Cells[3, 8].Value = "EM Task Step";
        //            ws.Cells[3, 9].Value = "Edit Type";
        //            ws.Cells[3, 10].Value = "Tool Group";
        //            ws.Cells[3, 11].Value = "Tool";
        //            ws.Cells[3, 12].Value = "Note";
        //            ws.Cells[3, 13].Value = "Revision Number";
        //            ws.Cells[3, 14].Value = "Revision Note";
        //            ws.Cells[3, 15].Value = "Revised Date";
        //            ws.Cells[3, 16].Value = "Revised by";
        //            ws.Cells[3, 17].Value = "Invalid Date";
        //            ws.Cells[4, 9].Value = "C";
        //            ws.Cells[5, 9].Value = "U";
        //            ws.Cells[6, 9].Value = "D";
        //            ws.Cells[4, 10].Value = "*";
        //            ws.Cells[5, 10].Value = "*";
        //            ws.Cells[6, 10].Value = "*";
        //            ws.Cells[4, 11].Value = "*";
        //            ws.Cells[5, 11].Value = "*";
        //            ws.Cells[6, 11].Value = "*";
        //            ws.Cells[4, 13].Value = "-";
        //            ws.Cells[5, 13].Value = "-";
        //            ws.Cells[6, 13].Value = "-";
        //            ws.Cells[4, 15].Value = "-";
        //            ws.Cells[5, 15].Value = "-";
        //            ws.Cells[6, 15].Value = "-";
        //            ws.Cells[4, 16].Value = "-";
        //            ws.Cells[5, 16].Value = "-";
        //            ws.Cells[6, 16].Value = "-";
        //            ws.Cells[4, 17].Value = "-";
        //            ws.Cells[5, 17].Value = "-";
        //            ws.Cells[6, 17].Value = "-";

        //            // ==== Formatting  ====
        //            allCells = ws.Cells["A1:Q6"];
        //            allCells.Style.Font.Name = "Calibri";
        //            allCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //            allCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //            ws.Row(1).Style.Font.Size = 20;
        //            ws.Row(2).Style.Font.Size = 11;
        //            ws.Row(3).Style.Font.Size = 10;
        //            ws.Row(4).Style.Font.Size = 11;
        //            ws.Row(5).Style.Font.Size = 11;
        //            ws.Row(6).Style.Font.Size = 11;

        //            ws.Cells[1, 1].Style.Font.Bold = true;
        //            ws.Cells["A1:B1"].Merge = true;
        //            ws.Cells[2, 1, 6, 17].AutoFitColumns();
        //            ws.Column(1).Width = 15;

        //            ws.Cells["A2:H3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A2:H3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 188, 188));

        //            ws.Cells["A1:H1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A1:H1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["A4:H6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A4:H6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["I2:Q3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["I2:Q3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 198, 220));

        //            ws.Cells["I4:Q6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["I4:Q6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(244, 245, 250));
        //            break;

        //        case "Part Master":
        //            ws.Cells[1, 1].Value = "Reference情報";
        //            ws.Cells[2, 1].Value = "ブロックNo";
        //            ws.Cells[2, 2].Value = "機種";
        //            ws.Cells[2, 3].Value = "Module名";
        //            ws.Cells[2, 4].Value = "OFF or ON";
        //            ws.Cells[2, 5].Value = "作業No";
        //            ws.Cells[2, 6].Value = "作業名称";
        //            ws.Cells[2, 7].Value = "EM Task";
        //            ws.Cells[2, 8].Value = "EM Task Step";
        //            ws.Cells[2, 9].Value = "ATA No";
        //            ws.Cells[2, 10].Value = "Fig. Item No";
        //            ws.Cells[2, 11].Value = "Edit Type";
        //            ws.Cells[2, 12].Value = "P / N";
        //            ws.Cells[2, 13].Value = "部品名称";
        //            ws.Cells[2, 14].Value = "100 % 部品フラグ";
        //            ws.Cells[2, 15].Value = "SB Non-Trackフラグ";
        //            ws.Cells[2, 16].Value = "PCC作業時間";
        //            ws.Cells[2, 17].Value = "洗浄種類";
        //            ws.Cells[2, 18].Value = "洗浄場所";
        //            ws.Cells[2, 19].Value = "エアブロー・ブラスト時間";
        //            ws.Cells[2, 20].Value = "占有サイズ";
        //            ws.Cells[2, 21].Value = "検査対象セクション";
        //            ws.Cells[2, 22].Value = "検査時間(VI)";
        //            ws.Cells[2, 23].Value = "検査時間(FPI)";
        //            ws.Cells[2, 24].Value = "検査時間(USI)";
        //            ws.Cells[2, 25].Value = "検査時間(ECI)";
        //            ws.Cells[2, 26].Value = "検査時間(MPI)";
        //            ws.Cells[2, 27].Value = "検査時間(CMM)";
        //            ws.Cells[2, 28].Value = "検査メモ";
        //            ws.Cells[2, 29].Value = "編成部品サイズ";
        //            ws.Cells[2, 30].Value = "KITBOX編成";
        //            ws.Cells[2, 31].Value = "OCR読取";
        //            ws.Cells[2, 32].Value = "Note";
        //            ws.Cells[2, 33].Value = "Revision Number";
        //            ws.Cells[2, 34].Value = "改訂理由";
        //            ws.Cells[2, 35].Value = "登録日";
        //            ws.Cells[2, 36].Value = "登録者";
        //            ws.Cells[2, 37].Value = "無効日";
        //            ws.Cells[3, 1].Value = "Block No";
        //            ws.Cells[3, 2].Value = "Eng Type";
        //            ws.Cells[3, 3].Value = "Module";
        //            ws.Cells[3, 4].Value = "OFF or ON";
        //            ws.Cells[3, 5].Value = "Operation No";
        //            ws.Cells[3, 6].Value = "Operation Name";
        //            ws.Cells[3, 7].Value = "EM Task";
        //            ws.Cells[3, 8].Value = "EM Task Step";
        //            ws.Cells[3, 9].Value = "ATA No";
        //            ws.Cells[3, 10].Value = "Fig. Item No";
        //            ws.Cells[3, 11].Value = "Edit Type";
        //            ws.Cells[3, 12].Value = "P/N";
        //            ws.Cells[3, 13].Value = "Part Name";
        //            ws.Cells[3, 14].Value = "100 % Replace";
        //            ws.Cells[3, 15].Value = "SB Non-Tracked";
        //            ws.Cells[3, 16].Value = "PCC working time";
        //            ws.Cells[3, 17].Value = "Cleaning Type";
        //            ws.Cells[3, 18].Value = "Cleaning Area";
        //            ws.Cells[3, 19].Value = "Air blow time";
        //            ws.Cells[3, 20].Value = "Size for cleaning";
        //            ws.Cells[3, 21].Value = "Insp section";
        //            ws.Cells[3, 22].Value = "VI working time";
        //            ws.Cells[3, 23].Value = "FPI working time";
        //            ws.Cells[3, 24].Value = "USI working time";
        //            ws.Cells[3, 25].Value = "ECI working time";
        //            ws.Cells[3, 26].Value = "MPI working time";
        //            ws.Cells[3, 27].Value = "CMM working time";
        //            ws.Cells[3, 28].Value = "Insp note";
        //            ws.Cells[3, 29].Value = "size for gathering";
        //            ws.Cells[3, 30].Value = "KIT BOX";
        //            ws.Cells[3, 31].Value = "OCR";
        //            ws.Cells[3, 32].Value = "Note";
        //            ws.Cells[3, 33].Value = "Revision Number";
        //            ws.Cells[3, 34].Value = "Revision Note";
        //            ws.Cells[3, 35].Value = "Revised Date";
        //            ws.Cells[3, 36].Value = "Revised by";
        //            ws.Cells[3, 37].Value = "Invalid Date";
        //            ws.Cells[4, 11].Value = "C";
        //            ws.Cells[5, 11].Value = "U";
        //            ws.Cells[6, 11].Value = "D";
        //            ws.Cells[4, 12].Value = "*";
        //            ws.Cells[5, 12].Value = "*";
        //            ws.Cells[6, 12].Value = "*";
        //            ws.Cells[4, 13].Value = "*";
        //            ws.Cells[4, 14].Value = "*";
        //            ws.Cells[4, 15].Value = "*";
        //            ws.Cells[4, 30].Value = "*";
        //            ws.Cells[4, 31].Value = "*";
        //            ws.Cells[4, 33].Value = "-";
        //            ws.Cells[5, 33].Value = "-";
        //            ws.Cells[6, 33].Value = "-";
        //            ws.Cells[4, 35].Value = "-";
        //            ws.Cells[5, 35].Value = "-";
        //            ws.Cells[6, 35].Value = "-";
        //            ws.Cells[4, 36].Value = "-";
        //            ws.Cells[5, 36].Value = "-";
        //            ws.Cells[6, 36].Value = "-";
        //            ws.Cells[4, 37].Value = "-";
        //            ws.Cells[5, 37].Value = "-";
        //            ws.Cells[6, 37].Value = "-";

        //            // ==== Formatting  ====
        //            allCells = ws.Cells["A1:AK6"];
        //            allCells.Style.Font.Name = "Calibri";
        //            allCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //            allCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //            ws.Row(1).Style.Font.Size = 20;
        //            ws.Row(2).Style.Font.Size = 11;
        //            ws.Row(3).Style.Font.Size = 10;
        //            ws.Row(4).Style.Font.Size = 11;
        //            ws.Row(5).Style.Font.Size = 11;
        //            ws.Row(6).Style.Font.Size = 11;

        //            ws.Cells[1, 1].Style.Font.Bold = true;
        //            ws.Cells["A1:B1"].Merge = true;
        //            ws.Cells[2, 1, 6, 37].AutoFitColumns();
        //            ws.Column(1).Width = 15;

        //            ws.Cells["A2:J3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A2:J3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 188, 188));

        //            ws.Cells["A1:J1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A1:J1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["A4:J6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A4:J6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["K2:AK3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["K2:AK3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 198, 220));

        //            ws.Cells["K4:AK6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["K4:AK6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(244, 245, 250));
        //            break;

        //        case "Parent and Child Master":
        //            ws.Cells[1, 1].Value = "Reference情報";
        //            ws.Cells[2, 1].Value = "ブロックNo";
        //            ws.Cells[2, 2].Value = "機種";
        //            ws.Cells[2, 3].Value = "Module名";
        //            ws.Cells[2, 4].Value = "OFF or ON";
        //            ws.Cells[2, 5].Value = "作業名称";
        //            ws.Cells[2, 6].Value = "Edit Type";
        //            ws.Cells[2, 7].Value = "作業No";
        //            ws.Cells[2, 8].Value = "親作業No";
        //            ws.Cells[2, 9].Value = "シーケンス";
        //            ws.Cells[2, 10].Value = "Note";
        //            ws.Cells[2, 11].Value = "改訂理由";
        //            ws.Cells[2, 12].Value = "登録日";
        //            ws.Cells[2, 13].Value = "登録者";
        //            ws.Cells[2, 14].Value = "更新日";
        //            ws.Cells[2, 15].Value = "更新者";
        //            ws.Cells[2, 16].Value = "無効日";
        //            ws.Cells[2, 17].Value = "無効者";
        //            ws.Cells[3, 1].Value = "Block No";
        //            ws.Cells[3, 2].Value = "Eng Type";
        //            ws.Cells[3, 3].Value = "Module";
        //            ws.Cells[3, 4].Value = "OFF or ON";
        //            ws.Cells[3, 5].Value = "Operation Name";
        //            ws.Cells[3, 6].Value = "Edit Type";
        //            ws.Cells[3, 7].Value = "Operation No";
        //            ws.Cells[3, 8].Value = "Parent Operation No";
        //            ws.Cells[3, 9].Value = "Seq";
        //            ws.Cells[3, 10].Value = "Note";
        //            ws.Cells[3, 11].Value = "Revision Note";
        //            ws.Cells[3, 12].Value = "Revised Date";
        //            ws.Cells[3, 13].Value = "Revised by";
        //            ws.Cells[3, 14].Value = "Updated Date";
        //            ws.Cells[3, 15].Value = "Updated by";
        //            ws.Cells[3, 16].Value = "Invalid Date";
        //            ws.Cells[3, 17].Value = "Invalid by";
        //            ws.Cells[4, 6].Value = "C";
        //            ws.Cells[5, 6].Value = "U";
        //            ws.Cells[6, 6].Value = "D";
        //            ws.Cells[4, 7].Value = "*";
        //            ws.Cells[5, 7].Value = "*";
        //            ws.Cells[6, 7].Value = "*";
        //            ws.Cells[4, 12].Value = "-";
        //            ws.Cells[5, 12].Value = "-";
        //            ws.Cells[6, 12].Value = "-";
        //            ws.Cells[4, 13].Value = "-";
        //            ws.Cells[5, 13].Value = "-";
        //            ws.Cells[6, 13].Value = "-";
        //            ws.Cells[4, 14].Value = "-";
        //            ws.Cells[5, 14].Value = "-";
        //            ws.Cells[6, 14].Value = "-";
        //            ws.Cells[4, 15].Value = "-";
        //            ws.Cells[5, 15].Value = "-";
        //            ws.Cells[6, 15].Value = "-";
        //            ws.Cells[4, 16].Value = "-";
        //            ws.Cells[5, 16].Value = "-";
        //            ws.Cells[6, 16].Value = "-";
        //            ws.Cells[4, 17].Value = "-";
        //            ws.Cells[5, 17].Value = "-";
        //            ws.Cells[6, 17].Value = "-";

        //            // ==== Formatting (EPPlus Style) ====
        //            allCells = ws.Cells["A1:Q6"];
        //            allCells.Style.Font.Name = "Calibri";
        //            allCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //            allCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //            ws.Row(1).Style.Font.Size = 20;
        //            ws.Row(2).Style.Font.Size = 11;
        //            ws.Row(3).Style.Font.Size = 10;
        //            ws.Row(4).Style.Font.Size = 11;
        //            ws.Row(5).Style.Font.Size = 11;
        //            ws.Row(6).Style.Font.Size = 11;

        //            ws.Cells[1, 1].Style.Font.Bold = true;
        //            ws.Cells["A1:B1"].Merge = true;
        //            ws.Cells[2, 1, 6, 17].AutoFitColumns();
        //            ws.Column(1).Width = 15;

        //            ws.Cells["A2:E3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A2:E3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 188, 188));

        //            ws.Cells["A1:E1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["A4:E6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A4:E6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["F2:Q3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["F2:Q3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 198, 220));

        //            ws.Cells["F4:Q6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["F4:Q6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(244, 245, 250));
        //            break;

        //        case "Tool Master":
        //            ws.Cells[1, 1].Value = "Reference情報";
        //            ws.Cells[2, 1].Value = "ブロックNo";
        //            ws.Cells[2, 2].Value = "機種";
        //            ws.Cells[2, 3].Value = "Module名";
        //            ws.Cells[2, 4].Value = "OFF or ON";
        //            ws.Cells[2, 5].Value = "作業No";
        //            ws.Cells[2, 6].Value = "作業名称";
        //            ws.Cells[2, 7].Value = "EM Task";
        //            ws.Cells[2, 8].Value = "EM Task Step";
        //            ws.Cells[2, 9].Value = "治工具グループ";
        //            ws.Cells[2, 10].Value = "Edit Type";
        //            ws.Cells[2, 11].Value = "治工具";
        //            ws.Cells[2, 12].Value = "治工具名称";
        //            ws.Cells[2, 13].Value = "Note";
        //            ws.Cells[2, 14].Value = "Revision Number";
        //            ws.Cells[2, 15].Value = "改訂理由";
        //            ws.Cells[2, 16].Value = "登録日";
        //            ws.Cells[2, 17].Value = "登録者";
        //            ws.Cells[2, 18].Value = "無効日";
        //            ws.Cells[3, 1].Value = "Block No";
        //            ws.Cells[3, 2].Value = "Eng Type";
        //            ws.Cells[3, 3].Value = "Module";
        //            ws.Cells[3, 4].Value = "OFF or ON";
        //            ws.Cells[3, 5].Value = "Operation No";
        //            ws.Cells[3, 6].Value = "Operation Name";
        //            ws.Cells[3, 7].Value = "EM Task";
        //            ws.Cells[3, 8].Value = "EM Task Step";
        //            ws.Cells[3, 9].Value = "Tool Group";
        //            ws.Cells[3, 10].Value = "Edit Type";
        //            ws.Cells[3, 11].Value = "Tool";
        //            ws.Cells[3, 12].Value = "Tool Name";
        //            ws.Cells[3, 13].Value = "Note";
        //            ws.Cells[3, 14].Value = "Revision Number";
        //            ws.Cells[3, 15].Value = "Revision Note";
        //            ws.Cells[3, 16].Value = "Revised Date";
        //            ws.Cells[3, 17].Value = "Revised by";
        //            ws.Cells[3, 18].Value = "Invalid Date";
        //            ws.Cells[4, 10].Value = "C";
        //            ws.Cells[5, 10].Value = "U";
        //            ws.Cells[6, 10].Value = "D";
        //            ws.Cells[4, 11].Value = "*";
        //            ws.Cells[5, 11].Value = "*";
        //            ws.Cells[6, 11].Value = "*";
        //            ws.Cells[4, 12].Value = "*";
        //            ws.Cells[4, 14].Value = "-";
        //            ws.Cells[5, 14].Value = "-";
        //            ws.Cells[6, 14].Value = "-";
        //            ws.Cells[4, 16].Value = "-";
        //            ws.Cells[5, 16].Value = "-";
        //            ws.Cells[6, 16].Value = "-";
        //            ws.Cells[4, 17].Value = "-";
        //            ws.Cells[5, 17].Value = "-";
        //            ws.Cells[6, 17].Value = "-";
        //            ws.Cells[4, 18].Value = "-";
        //            ws.Cells[5, 18].Value = "-";
        //            ws.Cells[6, 18].Value = "-";

        //            // ==== Formatting (EPPlus Style) ====
        //            allCells = ws.Cells["A1:R6"];
        //            allCells.Style.Font.Name = "Calibri";
        //            allCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //            allCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //            ws.Row(1).Style.Font.Size = 20;
        //            ws.Row(2).Style.Font.Size = 11;
        //            ws.Row(3).Style.Font.Size = 10;
        //            ws.Row(4).Style.Font.Size = 11;
        //            ws.Row(5).Style.Font.Size = 11;
        //            ws.Row(6).Style.Font.Size = 11;

        //            ws.Cells[1, 1].Style.Font.Bold = true;
        //            ws.Cells["A1:B1"].Merge = true;
        //            ws.Cells[2, 1, 6, 18].AutoFitColumns();
        //            ws.Column(1).Width = 15;

        //            ws.Cells["A2:F3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A2:F3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 188, 188));

        //            ws.Cells["A1:F1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["A4:F6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A4:F6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["G2:R3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["G2:R3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 198, 220));

        //            ws.Cells["G4:R6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["G4:R6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(244, 245, 250));
        //            break;


        //        case "EM Task and Part":
        //            ws.Cells[1, 1].Value = "Reference情報";
        //            ws.Cells[2, 1].Value = "ブロックNo";
        //            ws.Cells[2, 2].Value = "機種";
        //            ws.Cells[2, 3].Value = "Module名";
        //            ws.Cells[2, 4].Value = "OFF or ON";
        //            ws.Cells[2, 5].Value = "作業No";
        //            ws.Cells[2, 6].Value = "作業名称";
        //            ws.Cells[2, 7].Value = "Edit Type";
        //            ws.Cells[2, 8].Value = "EM Task";
        //            ws.Cells[2, 9].Value = "EM Task Step";
        //            ws.Cells[2, 10].Value = "ATA No";
        //            ws.Cells[2, 11].Value = "Fig. Item No";
        //            ws.Cells[2, 12].Value = "P/N";
        //            ws.Cells[2, 13].Value = "Description";
        //            ws.Cells[2, 14].Value = "QTY";
        //            ws.Cells[2, 15].Value = "Note";
        //            ws.Cells[2, 16].Value = "改訂理由";
        //            ws.Cells[2, 17].Value = "登録日";
        //            ws.Cells[2, 18].Value = "登録者";
        //            ws.Cells[2, 19].Value = "更新日";
        //            ws.Cells[2, 20].Value = "更新者";
        //            ws.Cells[2, 21].Value = "無効日";
        //            ws.Cells[2, 22].Value = "無効者";
        //            ws.Cells[3, 1].Value = "Block No";
        //            ws.Cells[3, 2].Value = "Eng Type";
        //            ws.Cells[3, 3].Value = "Module";
        //            ws.Cells[3, 4].Value = "OFF or ON";
        //            ws.Cells[3, 5].Value = "Operation No";
        //            ws.Cells[3, 6].Value = "Operation Name";
        //            ws.Cells[3, 7].Value = "Edit Type";
        //            ws.Cells[3, 8].Value = "EM Task";
        //            ws.Cells[3, 9].Value = "EM Task Step";
        //            ws.Cells[3, 10].Value = "ATA No";
        //            ws.Cells[3, 11].Value = "Fig. Item No";
        //            ws.Cells[3, 12].Value = "P/N";
        //            ws.Cells[3, 13].Value = "Description";
        //            ws.Cells[3, 14].Value = "QTY";
        //            ws.Cells[3, 15].Value = "Note";
        //            ws.Cells[3, 16].Value = "Revision Note";
        //            ws.Cells[3, 17].Value = "Revised Date";
        //            ws.Cells[3, 18].Value = "Revised by";
        //            ws.Cells[3, 19].Value = "Updated Date";
        //            ws.Cells[3, 20].Value = "Updated by";
        //            ws.Cells[3, 21].Value = "Invalid Date";
        //            ws.Cells[3, 22].Value = "Invalid by";
        //            ws.Cells[4, 7].Value = "C";
        //            ws.Cells[5, 7].Value = "U";
        //            ws.Cells[6, 7].Value = "D";
        //            ws.Cells[4, 8].Value = "*";
        //            ws.Cells[5, 8].Value = "*";
        //            ws.Cells[6, 8].Value = "*";
        //            ws.Cells[4, 9].Value = "*";
        //            ws.Cells[5, 9].Value = "*";
        //            ws.Cells[5, 9].Value = "*";
        //            ws.Cells[4, 10].Value = "*";
        //            ws.Cells[5, 10].Value = "*";
        //            ws.Cells[6, 10].Value = "*";
        //            ws.Cells[4, 11].Value = "*";
        //            ws.Cells[5, 11].Value = "*";
        //            ws.Cells[6, 11].Value = "*";
        //            ws.Cells[4, 12].Value = "*";
        //            ws.Cells[5, 12].Value = "*";
        //            ws.Cells[6, 12].Value = "*";
        //            ws.Cells[4, 13].Value = "*";
        //            ws.Cells[4, 14].Value = "*";
        //            ws.Cells[4, 17].Value = "-";
        //            ws.Cells[5, 17].Value = "-";
        //            ws.Cells[6, 17].Value = "-";
        //            ws.Cells[4, 18].Value = "-";
        //            ws.Cells[5, 18].Value = "-";
        //            ws.Cells[6, 18].Value = "-";
        //            ws.Cells[4, 19].Value = "-";
        //            ws.Cells[5, 19].Value = "-";
        //            ws.Cells[6, 19].Value = "-";
        //            ws.Cells[4, 20].Value = "-";
        //            ws.Cells[5, 20].Value = "-";
        //            ws.Cells[6, 20].Value = "-";
        //            ws.Cells[4, 21].Value = "-";
        //            ws.Cells[5, 21].Value = "-";
        //            ws.Cells[6, 21].Value = "-";
        //            ws.Cells[4, 22].Value = "-";
        //            ws.Cells[5, 22].Value = "-";
        //            ws.Cells[6, 22].Value = "-";

        //            // ==== Formatting (EPPlus Style) ====
        //            allCells = ws.Cells["A1:V6"];
        //            allCells.Style.Font.Name = "Calibri";
        //            allCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //            allCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //            // Font sizes
        //            ws.Row(1).Style.Font.Size = 20;
        //            ws.Row(2).Style.Font.Size = 11;
        //            ws.Row(3).Style.Font.Size = 10;
        //            ws.Row(4).Style.Font.Size = 11;
        //            ws.Row(5).Style.Font.Size = 11;
        //            ws.Row(6).Style.Font.Size = 11;

        //            // Bold cell A1
        //            ws.Cells[1, 1].Style.Font.Bold = true;
        //            ws.Cells["A1:B1"].Merge = true;
        //            ws.Cells[2, 1, 6, 22].AutoFitColumns();
        //            ws.Column(1).Width = 15;

        //            // Background colors
        //            ws.Cells["A2:F3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A2:F3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 188, 188));

        //            ws.Cells["A1:F1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["A4:F6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A4:F6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["G2:V3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["G2:V3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 198, 220));

        //            ws.Cells["G4:V6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["G4:V6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(244, 245, 250));
        //            break;


        //        case "EM Task and Tool Group":
        //            ws.Cells[1, 1].Value = "Reference情報";
        //            ws.Cells[2, 1].Value = "ブロックNo";
        //            ws.Cells[2, 2].Value = "機種";
        //            ws.Cells[2, 3].Value = "Module名";
        //            ws.Cells[2, 4].Value = "OFF or ON";
        //            ws.Cells[2, 5].Value = "作業No";
        //            ws.Cells[2, 6].Value = "作業名称";
        //            ws.Cells[2, 7].Value = "Edit Type";
        //            ws.Cells[2, 8].Value = "EM Task";
        //            ws.Cells[2, 9].Value = "EM Task Step"; ;
        //            ws.Cells[2, 10].Value = "治工具グループ";
        //            ws.Cells[2, 11].Value = "QTY";
        //            ws.Cells[2, 12].Value = "Note";
        //            ws.Cells[2, 13].Value = "改訂理由";
        //            ws.Cells[2, 14].Value = "登録日";
        //            ws.Cells[2, 15].Value = "登録者";
        //            ws.Cells[2, 16].Value = "更新日";
        //            ws.Cells[2, 17].Value = "更新者";
        //            ws.Cells[2, 18].Value = "無効日";
        //            ws.Cells[2, 19].Value = "無効者";
        //            ws.Cells[3, 1].Value = "Block No";
        //            ws.Cells[3, 2].Value = "Eng Type";
        //            ws.Cells[3, 3].Value = "Module";
        //            ws.Cells[3, 4].Value = "OFF or ON";
        //            ws.Cells[3, 5].Value = "Operation No";
        //            ws.Cells[3, 6].Value = "Operation Name";
        //            ws.Cells[3, 7].Value = "Edit Type";
        //            ws.Cells[3, 8].Value = "EM Task";
        //            ws.Cells[3, 9].Value = "EM Task Step";
        //            ws.Cells[3, 10].Value = "Tool Group";
        //            ws.Cells[3, 11].Value = "QTY";
        //            ws.Cells[3, 12].Value = "Note";
        //            ws.Cells[3, 13].Value = "Revision Note";
        //            ws.Cells[3, 14].Value = "Revised Date";
        //            ws.Cells[3, 15].Value = "Revised by";
        //            ws.Cells[3, 16].Value = "Updated Date";
        //            ws.Cells[3, 17].Value = "Updated by";
        //            ws.Cells[3, 18].Value = "Invalid Date";
        //            ws.Cells[3, 19].Value = "Invalid by";
        //            ws.Cells[4, 7].Value = "C";
        //            ws.Cells[5, 7].Value = "U";
        //            ws.Cells[6, 7].Value = "D";
        //            ws.Cells[4, 8].Value = "*";
        //            ws.Cells[5, 8].Value = "*";
        //            ws.Cells[6, 8].Value = "*";
        //            ws.Cells[4, 9].Value = "*";
        //            ws.Cells[5, 9].Value = "*";
        //            ws.Cells[6, 9].Value = "*";
        //            ws.Cells[4, 10].Value = "*";
        //            ws.Cells[5, 10].Value = "*";
        //            ws.Cells[6, 10].Value = "*";
        //            ws.Cells[4, 11].Value = "*";
        //            ws.Cells[4, 14].Value = "-";
        //            ws.Cells[5, 14].Value = "-";
        //            ws.Cells[6, 14].Value = "-";
        //            ws.Cells[4, 15].Value = "-";
        //            ws.Cells[5, 15].Value = "-";
        //            ws.Cells[6, 15].Value = "-";
        //            ws.Cells[4, 16].Value = "-";
        //            ws.Cells[5, 16].Value = "-";
        //            ws.Cells[6, 16].Value = "-";
        //            ws.Cells[4, 17].Value = "-";
        //            ws.Cells[5, 17].Value = "-";
        //            ws.Cells[6, 17].Value = "-";
        //            ws.Cells[4, 18].Value = "-";
        //            ws.Cells[5, 18].Value = "-";
        //            ws.Cells[6, 18].Value = "-";
        //            ws.Cells[4, 19].Value = "-";
        //            ws.Cells[5, 19].Value = "-";
        //            ws.Cells[6, 19].Value = "-";

        //            // ==== Formatting (EPPlus Style) ====
        //            allCells = ws.Cells["A1:S6"];
        //            allCells.Style.Font.Name = "Calibri";
        //            allCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //            allCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //            // Font sizes
        //            ws.Row(1).Style.Font.Size = 20;
        //            ws.Row(2).Style.Font.Size = 11;
        //            ws.Row(3).Style.Font.Size = 10;
        //            ws.Row(4).Style.Font.Size = 11;
        //            ws.Row(5).Style.Font.Size = 11;
        //            ws.Row(6).Style.Font.Size = 11;

        //            // Bold cell A1
        //            ws.Cells[1, 1].Style.Font.Bold = true;
        //            ws.Cells["A1:B1"].Merge = true;
        //            ws.Cells[2, 1, 6, 19].AutoFitColumns();
        //            ws.Column(1).Width = 15;

        //            // Background colors
        //            ws.Cells["A2:F3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A2:F3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 188, 188));

        //            ws.Cells["A1:F1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["A4:F6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["A4:F6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 218, 218));

        //            ws.Cells["G2:S3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["G2:S3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 198, 220));

        //            ws.Cells["G4:S6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            ws.Cells["G4:S6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(244, 245, 250));
        //            break;

        //        default:
        //            break;
        //    }

        //}

        private void btnParentChild_Click(object sender, EventArgs e)
        {
            var startTime = DateTime.Now;
            if (string.IsNullOrEmpty(txtOutputFolder.Text))
            {
                MessageBox.Show("Select a destination folder for the output files.");
                return;
            }

            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Select Parent Child master file");
                return;
            }

            if (string.IsNullOrEmpty(InputPath.Text))
            {
                MessageBox.Show("Select the directory that contains the source files.");
                return;
            }

            var sourceParentChildMaster = textBox1.Text;//"D:\\iHi\\LChart Inputs\\Batch-Deliverables\\Parent and Child Master.xlsx";

            var groupedList = new List<ItemGroup>();
            using (var package = new ExcelPackage(new FileInfo(sourceParentChildMaster)))
            {
                // Get the first worksheet
                var worksheetParentChildMaster = package.Workbook.Worksheets[1];

                // Find total rows and columns
                int worksheetParentChildMasterRowCount = worksheetParentChildMaster.Dimension.Rows;
                int worksheetParentChildMasterColumnCount = worksheetParentChildMaster.Dimension.Columns;

                Console.WriteLine($"Rows: {worksheetParentChildMasterRowCount}, Columns: {worksheetParentChildMasterColumnCount}");

                //Create new excel file
                if (File.Exists(outputFolder))
                {
                    File.Delete(outputFolder); // delete old file
                }

                try
                {
                    var rows = new List<ExcelRow>();

                    for (int referenceRow = 7; referenceRow <= worksheetParentChildMasterRowCount; referenceRow++)
                    {
                        rows.Add(new ExcelRow
                        {
                            BlockNumber = Convert.ToString(worksheetParentChildMaster.Cells[referenceRow, 1].Text),
                            Module = Convert.ToString(worksheetParentChildMaster.Cells[referenceRow, 3].Text),
                            Direction = Convert.ToString(worksheetParentChildMaster.Cells[referenceRow, 4].Text)
                        });
                    }

                    groupedList = rows
    .GroupBy(r => new { r.Module, r.Direction })
    .Select(g => new ItemGroup
    {
        ModuleName = $"{g.Key.Module} {g.Key.Direction}",
        Direction = g.Key.Direction,
        Blocks = g.Select(x => new BlockItem
        {
            BlockNumber = x.BlockNumber
        }).ToList()
    })
    .ToList();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            string[] files = Directory.GetFiles(InputPath.Text);

            try
            {
                foreach (var group in groupedList)
                {
                    if (!FilesAreSearchCompatible(group.ModuleName)) continue;

                    ExcelWorkbook workbookModule = null;
                    ExcelWorksheet workSheetLChart = null;
                    ExcelWorksheet workSheetManual = null;
                    ExcelWorksheet workSheetNovel = null;
                    ExcelWorksheet workSheetToTraverse = null;

                    var matchedFile = files
                   .Where(f => Path.GetFileName(f).StartsWith($"{group.ModuleName}", StringComparison.OrdinalIgnoreCase))
                   .FirstOrDefault();

                    if (matchedFile == null)
                    {
                        Console.WriteLine($"❌ No file found for: {group.ModuleName}");
                        continue;  // ← DO NOT STOP THE LOOP
                    }

                    try
                    {
                        var package = new ExcelPackage(new FileInfo(matchedFile));
                        workbookModule = package.Workbook;

                        var sheetNamesList = package.Workbook.Worksheets
                  .Select(ws => ws.Name)
                  .ToList();

                        foreach (var block in group.Blocks)
                        {
                            if (string.IsNullOrEmpty(block.BlockNumber))
                            {
                                continue;
                            }

                            bool isNovelBlock = block.BlockNumber.IndexOf("NOVEL", StringComparison.OrdinalIgnoreCase) >= 0;

                            // Sheet references
                            workSheetLChart = package.Workbook.Worksheets[SheetLChart];

                            workSheetNovel = sheetNamesList.Contains(SheetLChartNovel)
                               ? package.Workbook.Worksheets[SheetLChartNovel]
                               : null;

                            bool found = false;
                            int foundAtRow = 0;
                            int foundAtColumn = 0;

                            // ============================================================
                            // CASE 1: NOVEL BLOCK → Search ONLY Novel sheet
                            // ============================================================
                            if (isNovelBlock)
                            {
                                if (workSheetNovel == null)
                                    continue; // skip novel blocks with no novel sheet

                                if (SearchBlockInSheet(workSheetNovel, block.BlockNumber, out foundAtRow, out foundAtColumn))
                                {
                                    found = true;
                                    workSheetToTraverse = workSheetNovel;
                                }
                                else
                                {
                                    continue; // not found
                                }
                            }
                            else
                            {
                                // ============================================================
                                // CASE 2: NORMAL BLOCK → Search L-Chart first, then Novel
                                // ============================================================

                                // Try main sheet
                                if (SearchBlockInSheet(workSheetLChart, block.BlockNumber, out foundAtRow, out foundAtColumn))
                                {
                                    found = true;
                                    workSheetToTraverse = workSheetLChart;
                                }
                                else if (workSheetNovel != null &&
                                         SearchBlockInSheet(workSheetNovel, block.BlockNumber, out foundAtRow, out foundAtColumn))
                                {
                                    found = true;
                                    workSheetToTraverse = workSheetNovel;
                                }
                                else
                                {
                                    continue; // not found anywhere
                                }
                            }

                            if (found)
                            {
                                workSheetManual = workbookModule.Worksheets[SheetManual];

                                if (group.Direction == "ON")
                                {
                                    var downLineStartsAtRow = foundAtRow + 4;
                                    var downLineStartsAtColumn = foundAtColumn - 2;
                                    var parentBlocks = TraverseDown(downLineStartsAtRow, downLineStartsAtColumn, workSheetToTraverse);
                                    
                                    foreach (var p in parentBlocks)
                                    {
                                        var operationNumber = ReadOperationNoFromManualSheet(workSheetManual, p.Text);
                                        // ⭐ Add the ParentInfo
                                        block.Parents.Add(new ParentInfo
                                        {
                                            ParentNumber = p.Text,
                                            ParentOperationNumber = operationNumber
                                        });
                                    }
                                }
                                else if (group.Direction == "OFF")
                                {
                                    var upLineStartsAtRow = foundAtRow - 1;
                                    var upLineStartsAtColumn = foundAtColumn - 2;
                                    var parentBlocks = TraverseUp(upLineStartsAtRow, upLineStartsAtColumn, workSheetToTraverse);
                                    
                                    foreach (var p in parentBlocks)
                                    {
                                        var operationNumber = ReadOperationNoFromManualSheet(workSheetManual, p.Text);
                                        // ⭐ Add the ParentInfo
                                        block.Parents.Add(new ParentInfo
                                        {
                                            ParentNumber = p.Text,
                                            ParentOperationNumber = operationNumber
                                        });
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("ERROR: " + ex.Message);
                    }
                    finally
                    {
                    }
                }

                // After processing all groups, create output Excel file and update its H Column (Parent Operation No)

                string copiedFile = Path.Combine(outputFolder, "newFile.xlsx");

                // Delete if already exists
                if (File.Exists(copiedFile))
                {
                    File.Delete(copiedFile);
                }

                // Copy the new file
                File.Copy(sourceParentChildMaster, copiedFile);

                using (var package = new ExcelPackage(new FileInfo(copiedFile)))
                {
                    // Get the first worksheet
                    var worksheetOfCopiedFile = package.Workbook.Worksheets[1];
                    int lastRow = worksheetOfCopiedFile.Dimension.End.Row;
                    int lastCol = worksheetOfCopiedFile.Dimension.End.Column;

                    foreach (var group in groupedList)
                    {
                        var moduleName = group.ModuleName;

                        foreach (var block in group.Blocks)
                        {
                            UpdateExcelForBlock(worksheetOfCopiedFile, block, moduleName, ref lastRow, lastCol);
                        }
                    }
                    package.Save();
                }
                // end
            }
            finally
            {
            }
            var endTime = DateTime.Now;
            Console.WriteLine($"Start Time : {startTime}");
            Console.WriteLine($"End Time   : {endTime}");
            Console.WriteLine($"Total Time : {endTime - startTime}");
        }

        private bool FilesAreSearchCompatible(string module)
        {
            if (module == "CIC ON" || module == "FINAL ON" || module == "GBX ON" || module == "LPT_TEC OFF")
            {
                return false;
            }
            return true;
        }

        private string ReadOperationNoFromManualSheet(ExcelWorksheet manualWorkSheet, string parent)
        {
            bool operationNumberFound = false;
            int operationNumberfoundAtRow = 0;
            int operationNumberFoundAtColumn = 0;

            int lastRow = manualWorkSheet.Dimension?.End.Row ?? 1;

            for (int rrow = 1; rrow <= lastRow && !operationNumberFound; rrow++)
            {
                var cellText = manualWorkSheet.Cells[rrow, 1].Text?.Trim();
                cellText = cellText.Replace("\r", "")
           .Replace("\n", "")
           .Trim();

                cellText = Convert.ToString(cellText);

                if (string.Equals(cellText, parent, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"✅ Found \"{parent}\" at Row: {rrow}, Column: {1}");
                    operationNumberFound = true;
                    operationNumberfoundAtRow = rrow;
                    operationNumberFoundAtColumn = 1;
                    break;
                }
            }

            string operationNumber = "";
            if (operationNumberFound)
            {
                operationNumber = manualWorkSheet.Cells[operationNumberfoundAtRow, 7].Text;
            }
            return operationNumber;
        }

        public List<ExcelRangeBase> TraverseDown(int startRow, int startColumn, ExcelWorksheet _ws)
        {
            var result = new List<ExcelRangeBase>();
            DownQueue.Enqueue((startRow, startColumn));

            while (DownQueue.Count > 0)
            {
                var node = DownQueue.Dequeue();
                int totalRows = _ws.Dimension.End.Row;
                ProcessDown(node.row, node.col, _ws, totalRows, result);
            }
            return result;
        }

        // =======================================================
        //  MAIN ENTRY POINT
        // =======================================================
        public List<ExcelRangeBase> TraverseUp(int startRow, int startColumn, ExcelWorksheet _ws)
        {
            var result = new List<ExcelRangeBase>();
            UpQueue.Enqueue((startRow, startColumn));

            while (UpQueue.Count > 0)
            {
                var node = UpQueue.Dequeue();
                ProcessUp(node.row, node.col, _ws, result);
            }
            return result;
        }

        // =======================================================
        //  UP NAVIGATION
        // =======================================================
        private void ProcessUp(int row, int column, ExcelWorksheet _ws, List<ExcelRangeBase> result)
        {
            int r = row;

            while (r > 1)
            {
                ExcelRangeBase leftCell = _ws.Cells[r, column - 1];
                ExcelRangeBase rightCell = _ws.Cells[r, column];

                bool leftHasTop = HasTop(leftCell);
                bool rightHasTop = HasTop(rightCell);

                bool leftHasRight = HasRight(leftCell);
                bool rightHasLeft = HasLeft(rightCell);

                // -----------------------------------------
                // STOP condition:
                // left has NO right border AND right has NO left border
                // -----------------------------------------
                if (!leftHasRight && !rightHasLeft)
                    return;

                // -----------------------------------------
                // FOUND PARENT MERGED CELL
                // -----------------------------------------
                if (leftHasTop && rightHasTop)
                {
                    ExcelRangeBase parent;
                    if (TryGetImmediateMergeParent(leftCell, _ws, out parent))
                    {
                        result.Add(parent);
                        return;
                    }
                }
                if (leftHasTop)
                {
                    var turnLeft = (ExcelRangeBase)leftCell.Offset(-1, 0);
                    ProcessLeftPath(turnLeft);
                }

                if (rightHasTop)
                {
                    var turnRight = (ExcelRangeBase)rightCell.Offset(-1, 0);
                    ProcessRightPath(turnRight);
                }

                r--; // MOVE UP
            }
        }

        // =======================================================
        //  LEFT PATH
        // =======================================================
        private void ProcessLeftPath(ExcelRangeBase startCell)
        {
            //ExcelRange current = startCell.Offset[-1, 0];
            var current = startCell;//.Offset[-1, 0];

            while (true)
            {
                bool left = HasLeft(current);
                bool bottom = HasBottom(current);

                if (!bottom)
                    break;

                if (left && bottom)
                {
                    // Enqueue NEW UP traversal point
                    UpQueue.Enqueue((current.Start.Row, current.Start.Column));
                }

                current = (ExcelRangeBase)current.Offset(0, -1); // MOVE LEFT
            }
        }

        // =======================================================
        //  RIGHT PATH
        // =======================================================
        private void ProcessRightPath(ExcelRangeBase startCell)
        {
            var current = startCell;

            while (true)
            {
                bool right = HasRight(current);
                bool bottom = HasBottom(current);

                if (!bottom)
                    break;

                if (right && bottom)
                {
                    // Enqueue NEW UP traversal point
                    int move1ColumnRight = current.Start.Column + 1;
                    UpQueue.Enqueue((current.Start.Row, move1ColumnRight));
                }

                current = (ExcelRangeBase)current.Offset(0, 10); // MOVE RIGHT
            }
        }

        // =======================================================
        //  RESOLVE MERGED CELL PARENT
        // =======================================================
        //private Excel.Range ResolveMergeParent(Excel.Range belowCell, Worksheet _ws)
        //{
        //    Excel.Range above = belowCell.Offset[-2, 0];

        //    if (!above.MergeCells)
        //        return null;

        //    Excel.Range merged = above.MergeArea;

        //    Excel.Range topLeft = merged.Cells[1, 1];
        //    int parentCol = topLeft.Column + merged.Columns.Count;

        //    return _ws.Cells[topLeft.Row, parentCol];
        //}

        // =======================================================
        //  BORDER HELPERS
        // =======================================================
        private ExcelRangeBase GetActualCell(ExcelRangeBase cell)
        {
            if (cell.Merge)
            {
                string merged = cell.Worksheet.MergedCells[cell.Start.Row, cell.Start.Column];
                var addr = new ExcelAddress(merged);

                // Returns ExcelRangeBase (safe)
                return cell.Worksheet.Cells[addr.Start.Row, addr.Start.Column];
            }

            return cell;
        }

        private bool HasTop(ExcelRangeBase c)
        {
            var cell = GetActualCell(c);
            var ws = cell.Worksheet;

            int row = cell.Start.Row;
            int col = cell.Start.Column;

            // 1. Direct top border on current cell
            if (cell.Style.Border.Top.Style != ExcelBorderStyle.None)
                return true;

            // 2. Look at the cell above
            if (ws.Dimension != null && row > 1)
            {
                int aboveRow = row - 1;
                var above = ws.Cells[aboveRow, col];

                // Check merged region above
                var mergedAddress = ws.MergedCells[aboveRow, col];
                if (!string.IsNullOrEmpty(mergedAddress))
                {
                    var addr = new ExcelAddress(mergedAddress);

                    // Use merged region's FIRST ROW but SAME COLUMN
                    above = ws.Cells[addr.Start.Row, col];
                }

                // The bottom border of the cell above forms our visible top border
                if (above.Style.Border.Bottom.Style != ExcelBorderStyle.None)
                    return true;
            }

            return false;
        }

        private bool HasBottom(ExcelRangeBase c)
        {
            var cell = GetActualCell(c);
            var ws = cell.Worksheet;

            int row = cell.Start.Row;
            int col = cell.Start.Column;

            // 1. Direct bottom border on this cell
            if (cell.Style.Border.Bottom.Style != ExcelBorderStyle.None)
                return true;

            // 2. Look at the cell below
            if (ws.Dimension != null && row < ws.Dimension.End.Row)
            {
                int nextRow = row + 1;
                var below = ws.Cells[nextRow, col];

                // Check merged region below
                var mergedAddress = ws.MergedCells[nextRow, col];
                if (!string.IsNullOrEmpty(mergedAddress))
                {
                    var addr = new ExcelAddress(mergedAddress);

                    // Use merged region's first row but same column
                    below = ws.Cells[addr.Start.Row, col];
                }

                // Now check the visible top border
                if (below.Style.Border.Top.Style != ExcelBorderStyle.None)
                    return true;
            }

            return false;
        }


        private bool HasLeft(ExcelRangeBase c)
        {
            var cell = GetActualCell(c);
            return cell.Style.Border.Left.Style != ExcelBorderStyle.None;
        }

        private bool HasRight(ExcelRangeBase c)
        {
            var cell = GetActualCell(c);
            return cell.Style.Border.Right.Style != ExcelBorderStyle.None;
        }



        private bool TryGetImmediateMergeParent(ExcelRangeBase belowCell, ExcelWorksheet ws, out ExcelRangeBase parentCell)
        {
            parentCell = null;

            // Move 2 rows up from belowCell
            int targetRow = belowCell.Start.Row - 2;
            int targetCol = belowCell.Start.Column;
            ExcelRangeBase above = ws.Cells[targetRow, targetCol];

            // Not merged → no parent
            if (!above.Merge)
                return false;

            // Find the merged range that contains 'above'
            string mergedAddress = ws.MergedCells
                .FirstOrDefault(a =>
                {
                    var addr = new ExcelAddress(a);
                    return above.Start.Row >= addr.Start.Row && above.Start.Row <= addr.End.Row
                        && above.Start.Column >= addr.Start.Column && above.Start.Column <= addr.End.Column;
                });

            if (mergedAddress == null)
                return false;

            // Full merged range
            ExcelRangeBase merged = ws.Cells[mergedAddress];

            // Top-left of merged block
            int topLeftRow = merged.Start.Row;
            int topLeftCol = merged.Start.Column;

            // Compute parent column = right edge + 1
            int parentCol = topLeftCol + merged.Columns;

            // Get parent cell
            parentCell = ws.Cells[topLeftRow, parentCol];

            return true;
        }


        // =======================================================
        //  UP NAVIGATION
        // =======================================================
        private void ProcessDown(int row, int column, ExcelWorksheet _ws, int worksheetRowsCount, List<ExcelRangeBase> result)
        {
            int r = row;

            while (r <= worksheetRowsCount)
            {
                ExcelRangeBase leftCell = _ws.Cells[r, column - 1];
                ExcelRangeBase rightCell = _ws.Cells[r, column];

                bool leftHasBottom = HasBottom(leftCell);
                bool rightHasBottom = HasBottom(rightCell);

                bool leftHasRight = HasRight(leftCell);
                bool rightHasLeft = HasLeft(rightCell);

                // -----------------------------------------
                // STOP condition:
                // left has NO right border AND right has NO left border
                // -----------------------------------------
                if (!leftHasRight && !rightHasLeft)
                    return;

                // -----------------------------------------
                // FOUND PARENT MERGED CELL
                // -----------------------------------------
                if (leftHasBottom && rightHasBottom)
                {
                    ExcelRangeBase parent;
                    if (TryGetImmediateMergeParentDown(leftCell, _ws, out parent))
                    {
                        result.Add(parent);
                        return;
                    }
                }
                if (leftHasBottom)
                {
                    var turnLeft = (ExcelRangeBase)leftCell.Offset(1, 0); // Move 1 row down
                    ProcessDownLeftPath(turnLeft);
                }

                if (rightHasBottom)
                {
                    var turnRight = (ExcelRangeBase)rightCell.Offset(1, 0); // Move 1 row down
                    ProcessDownRightPath(turnRight);
                }

                r++; // MOVE DOWN
            }
        }

        // =======================================================
        //  LEFT PATH
        // =======================================================
        private void ProcessDownLeftPath(ExcelRangeBase startCell)
        {
            //ExcelRange current = startCell.Offset[-1, 0];
            ExcelRangeBase current = startCell;//.Offset[-1, 0];

            while (true)
            {
                bool left = HasLeft(current);
                bool top = HasTop(current);

                if (!top)
                    break;

                if (left && top)
                {
                    // Enqueue NEW UP traversal point
                    int row = current.Start.Row;
                    int column = current.Start.Column;
                    DownQueue.Enqueue((row, column));
                }

                current = (ExcelRangeBase)current.Offset(0, -1); // MOVE LEFT
            }
        }

        // =======================================================
        //  RIGHT PATH
        // =======================================================
        private void ProcessDownRightPath(ExcelRangeBase startCell)
        {
            ExcelRangeBase current = startCell;

            while (true)
            {
                bool right = HasRight(current);
                bool top = HasTop(current);

                if (!top)
                    break;

                if (right && top)
                {
                    // Enqueue NEW UP traversal point
                    int move1ColumnRight = current.Start.Column + 1;
                    DownQueue.Enqueue((current.Start.Row, move1ColumnRight));
                }

                current = (ExcelRangeBase)current.Offset(0, 1); // Move 1 column right
            }
        }

        private bool TryGetImmediateMergeParentDown(ExcelRangeBase belowCell, ExcelWorksheet ws, out ExcelRangeBase parentCell)
        {
            parentCell = null;

            // Move 2 rows down from belowCell (Interop Offset[-2,0])
            int targetRow = belowCell.Start.Row + 2;
            int targetCol = belowCell.Start.Column;
            ExcelRangeBase above = ws.Cells[targetRow, targetCol];

            // Not merged → no parent
            if (!above.Merge)
                return false;

            // Find the merged range that contains 'above'
            string mergedAddress = ws.MergedCells
                .FirstOrDefault(a =>
                {
                    var addr = new ExcelAddress(a);
                    return above.Start.Row >= addr.Start.Row && above.Start.Row <= addr.End.Row
                        && above.Start.Column >= addr.Start.Column && above.Start.Column <= addr.End.Column;
                });

            if (mergedAddress == null)
                return false;

            // Full merged range
            ExcelRangeBase merged = ws.Cells[mergedAddress];

            // Top-left of merged block
            int topLeftRow = merged.Start.Row;
            int topLeftCol = merged.Start.Column;

            // Compute parent column = right edge + 1
            int parentCol = topLeftCol + merged.Columns;

            // Get parent cell
            parentCell = ws.Cells[topLeftRow, parentCol];

            return true;
        }


        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select a file";
                ofd.Filter = "All Files (*.*)|*.*";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    //txtFilePath.Text = ofd.FileName;
                }
            }
        }

        private void WriteToNewFile(int writeAtRow, object[] columnValues)
        {
            using (var package = new ExcelPackage(new FileInfo(outputFolder)))
            {
                // Ensure at least one worksheet exists
                var ws = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add("Sheet1");

                for (int i = 0; i < columnValues.Length; i++)
                {
                    ws.Cells[writeAtRow, i + 1].Value = columnValues[i];
                }
                package.Save();
            }
        }

        private void btnOutputPathBrowse_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            dialog.InitialDirectory = "C:\\";

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string selectedPath = dialog.FileName;
                //MessageBox.Show("Selected Folder: " + selectedPath);
                txtOutputFolder.Text = selectedPath;
                outputFolder = selectedPath;
            }
        }

        /// <summary>
        /// Updates Excel sheet for a given BlockItem and its parents.
        /// Inserts rows if multiple parents exist.
        /// </summary>
        /// <param name="ws">The Excel worksheet</param>
        /// <param name="block">The BlockItem</param>
        /// <param name="moduleName">Module name from the group</param>
        /// <param name="direction">Direction from the group</param>
        /// <param name="startRow">Row to start searching from (usually 2)</param>
        private static void UpdateExcelForBlock(ExcelWorksheet ws, BlockItem block,
            string moduleName, ref int lastRow, int lastCol)
        {
            string blockNumber = block.BlockNumber?.Trim();
            var parents = block.Parents;

            if (parents == null || parents.Count == 0)
            {
                Console.WriteLine($"No parents found for Block {blockNumber} in Module {moduleName}");
                return;
            }

            // Find the matching row
            int row = -1;
            for (int r = 7; r <= lastRow; r++)
            {
                string excelBlock = ws.Cells[r, 1].Text.Trim();
                string excelModule = ws.Cells[r, 3].Text.Trim();
                string excelDirection = ws.Cells[r, 4].Text.Trim();

                if (excelBlock == blockNumber && $"{excelModule} {excelDirection}".Trim().Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                {
                    row = r;
                    break;
                }
            }

            if (row == -1)
            {
                Console.WriteLine($"Row not found for Block={blockNumber}, Module={moduleName}");
                return;
            }

            // -----------------------
            // Single parent
            // -----------------------
            if (parents.Count == 1)
            {
                ws.Cells[row, 8].Value = parents[0].ParentOperationNumber;
                return;
            }

            // -----------------------
            // Multiple parents → insert extra rows
            // -----------------------
            int extraRows = parents.Count - 1;
            ws.InsertRow(row + 1, extraRows, row); // copies formatting

            // Copy all column values from original row to inserted rows (except column H)
            for (int i = 1; i < parents.Count; i++)
            {
                int newRow = row + i;
                for (int col = 1; col <= lastCol; col++)
                {
                    if (col == 8) continue; // skip ParentOperationNumber column
                    ws.Cells[newRow, col].Value = ws.Cells[row, col].Value;
                }
            }

            // Write ParentOperationNumber for all parents
            for (int i = 0; i < parents.Count; i++)
            {
                int targetRow = row + i;
                ws.Cells[targetRow, 8].Value = parents[i].ParentOperationNumber;
            }

            // Update lastRow because we inserted rows
            lastRow += extraRows;
        }

        private bool SearchBlockInSheet(ExcelWorksheet ws, string blockNumber,
                                out int foundAtRow, out int foundAtColumn)
        {
            foundAtRow = 0;
            foundAtColumn = 0;

            if (ws?.Dimension == null)
                return false;

            int lastRow = ws.Dimension.End.Row;
            int lastCol = ws.Dimension.End.Column;

            for (int r = 1; r <= lastRow; r++)
            {
                for (int c = 1; c <= lastCol; c++)
                {
                    string text = ws.Cells[r, c].Text?
                        .Replace("\r", "")
                        .Replace("\n", "")
                        .Trim();

                    if (string.Equals(text, blockNumber, StringComparison.OrdinalIgnoreCase))
                    {
                        foundAtRow = r;
                        foundAtColumn = c;
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
