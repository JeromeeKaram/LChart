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
        const string SheetLChartHPTON = "L-CHART(POST0157)";
        const string HPTON = "HPT ON";

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
                    txtPortalFilePath.Text = selectedPath;
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
                        txtPortalFilePath.Text = fileDlg.FileName;
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
            string path = txtPortalFilePath.Text.Trim();
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
                string portalPath = txtPortalFilePath.Text.Trim(); string inputFolder = InputPath.Text.Trim(); string excelName = comboBox1.Text.Trim();

                if (string.IsNullOrWhiteSpace(inputFolder) && string.IsNullOrWhiteSpace(portalPath))
                {
                    MessageBox.Show("Please select an input folder and LChart portal file paths.",
                                    "Missing Path",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    txtPortalFilePath.Focus();
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
                    txtPortalFilePath.Focus();
                    return;
                }

                if (!System.IO.File.Exists(portalPath) && !System.IO.Directory.Exists(portalPath))
                {
                    MessageBox.Show("The selected portal file or folder path does not exist.",
                                    "Invalid Path",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    txtPortalFilePath.Focus();
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

        private void btnParentChild_Click(object sender, EventArgs e)
        {
            var startTime = DateTime.Now;
            if (string.IsNullOrEmpty(txtOutputFolder.Text))
            {
                MessageBox.Show("Select a destination folder for the output files.");
                return;
            }

            if (string.IsNullOrEmpty(txtPortalFilePath.Text))
            {
                MessageBox.Show("Select Parent Child master file");
                return;
            }

            if (string.IsNullOrEmpty(InputPath.Text))
            {
                MessageBox.Show("Select the directory that contains the source files.");
                return;
            }

            var sourceParentChildMaster = txtPortalFilePath.Text;//"D:\\iHi\\LChart Inputs\\Batch-Deliverables\\Parent and Child Master.xlsx";

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

                    //if (group.ModuleName!= "CIC OFF") continue;

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

                        if (group.ModuleName.IndexOf(HPTON, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            workSheetLChart = package.Workbook.Worksheets[SheetLChartHPTON];
                        }
                        else
                        {
                            workSheetLChart = package.Workbook.Worksheets[SheetLChart];
                        }

                        workSheetNovel = sheetNamesList.Contains(SheetLChartNovel)
                           ? package.Workbook.Worksheets[SheetLChartNovel]
                           : null;

                        foreach (var block in group.Blocks)
                        {
                            if (string.IsNullOrEmpty(block.BlockNumber))
                            {
                                continue;
                            }

                            bool isNovelBlock = block.BlockNumber.IndexOf("NOVEL", StringComparison.OrdinalIgnoreCase) >= 0;

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
                                        var leftCell = workSheetToTraverse.Cells[p.Start.Row, p.Start.Column - 1];

                                        string mergedCellText = null;

                                        if (leftCell.Merge)
                                        {
                                            // Get merged range address (e.g. "A3:A6")
                                            var mergedAddress = workSheetToTraverse.MergedCells[leftCell.Start.Row, leftCell.Start.Column];

                                            // Get the merged range
                                            var mergedRange = workSheetToTraverse.Cells[mergedAddress];

                                            // TOP-LEFT cell of merged range
                                            var topLeftCell = workSheetToTraverse.Cells[
                                                mergedRange.Start.Row,
                                                mergedRange.Start.Column
                                            ];

                                            mergedCellText = topLeftCell.Value?.ToString().Trim();

                                            bool isDummy = false;

                                            if (string.Equals(mergedCellText, "Dummy", StringComparison.OrdinalIgnoreCase))
                                            {
                                                isDummy = true;
                                            }

                                            var operationNumber = ReadOperationNoFromManualSheet(workSheetManual, p.Text, isDummy);

                                            // ⭐ Add the ParentInfo
                                            block.Parents.Add(new ParentInfo
                                            {
                                                ParentNumber = p.Text,
                                                ParentOperationNumber = operationNumber
                                            });

                                        }
                                    }
                                }
                                else if (group.Direction == "OFF")
                                {
                                    var upLineStartsAtRow = foundAtRow - 1;
                                    var upLineStartsAtColumn = foundAtColumn - 2;
                                    var parentBlocks = TraverseUp(upLineStartsAtRow, upLineStartsAtColumn, workSheetToTraverse);

                                    foreach (var p in parentBlocks)
                                    {
                                        var leftCell = workSheetToTraverse.Cells[p.Start.Row, p.Start.Column - 1];

                                        string mergedCellText = null;

                                        if (leftCell.Merge)
                                        {
                                            // Get merged range address (e.g. "A3:A6")
                                            var mergedAddress = workSheetToTraverse.MergedCells[leftCell.Start.Row, leftCell.Start.Column];

                                            // Get the merged range
                                            var mergedRange = workSheetToTraverse.Cells[mergedAddress];

                                            // TOP-LEFT cell of merged range
                                            var topLeftCell = workSheetToTraverse.Cells[
                                                mergedRange.Start.Row,
                                                mergedRange.Start.Column
                                            ];

                                            mergedCellText = topLeftCell.Value?.ToString().Trim();

                                            bool isDummy =false;

                                            if (string.Equals(mergedCellText, "Dummy", StringComparison.OrdinalIgnoreCase))
                                            {
                                                isDummy = true;
                                            }
                                                
                                            var operationNumber = ReadOperationNoFromManualSheet(workSheetManual, p.Text, isDummy);

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
            MessageBox.Show($"Finished L-Chart processing in {endTime - startTime}");
        }

        private bool FilesAreSearchCompatible(string module)
        {
            if (module == "CIC ON" || module == "FINAL ON" || module == "GBX ON" || module == "LPT_TEC OFF")
            {
                return false;
            }
            return true;
        }

        private string ReadOperationNoFromManualSheet(ExcelWorksheet manualWorkSheet, string parent, bool isDummy)
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

            if (isDummy) 
            {
                operationNumber = "Dummy";
            } 
                
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
