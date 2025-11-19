using System;
using System.Windows.Forms;

namespace LChart_Comparison_Tool
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
            this.Help = new System.Windows.Forms.TabPage();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.linkLabel3 = new System.Windows.Forms.LinkLabel();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.label3 = new System.Windows.Forms.Label();
            this.Helpbtn = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.pictureBox5 = new System.Windows.Forms.PictureBox();
            this.LChart_Comparison = new System.Windows.Forms.TabPage();
            this.label16 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.InputBrowse = new System.Windows.Forms.Button();
            this.InputPath = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.BrowseBtn = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.Form1_Update = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.Tabs = new System.Windows.Forms.TabControl();
            this.Version_Control = new System.Windows.Forms.TabPage();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.btnParentChild = new System.Windows.Forms.Button();
            this.Help.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).BeginInit();
            this.LChart_Comparison.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.Tabs.SuspendLayout();
            this.Version_Control.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // Help
            // 
            this.Help.AutoScroll = true;
            this.Help.BackColor = System.Drawing.Color.Transparent;
            this.Help.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Help.Controls.Add(this.label8);
            this.Help.Controls.Add(this.label7);
            this.Help.Controls.Add(this.label4);
            this.Help.Controls.Add(this.linkLabel3);
            this.Help.Controls.Add(this.linkLabel2);
            this.Help.Controls.Add(this.linkLabel1);
            this.Help.Controls.Add(this.label3);
            this.Help.Controls.Add(this.Helpbtn);
            this.Help.Controls.Add(this.label2);
            this.Help.Controls.Add(this.pictureBox5);
            this.Help.Location = new System.Drawing.Point(4, 33);
            this.Help.Name = "Help";
            this.Help.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.Help.Size = new System.Drawing.Size(776, 320);
            this.Help.TabIndex = 4;
            this.Help.Text = "   Help  ";
            this.Help.Paint += new System.Windows.Forms.PaintEventHandler(this.Form1_Paint);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Yu Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.label8.Location = new System.Drawing.Point(102, 156);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(70, 16);
            this.label8.TabIndex = 24;
            this.label8.Text = "Developed:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Yu Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.label7.Location = new System.Drawing.Point(102, 116);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(64, 16);
            this.label7.TabIndex = 23;
            this.label7.Text = "Reviewed:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Yu Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.label4.Location = new System.Drawing.Point(102, 72);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(66, 16);
            this.label4.TabIndex = 22;
            this.label4.Text = "Innovated:";
            // 
            // linkLabel3
            // 
            this.linkLabel3.AutoSize = true;
            this.linkLabel3.Location = new System.Drawing.Point(246, 155);
            this.linkLabel3.Name = "linkLabel3";
            this.linkLabel3.Size = new System.Drawing.Size(346, 20);
            this.linkLabel3.TabIndex = 21;
            this.linkLabel3.TabStop = true;
            this.linkLabel3.Text = "Sudhakar Vemuri (Sudhakar.Vemuri@cyient.com)";
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Location = new System.Drawing.Point(246, 112);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(300, 20);
            this.linkLabel2.TabIndex = 20;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "Naveen Korivela (korivela1022@ihi-g.com)";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(246, 69);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(307, 20);
            this.linkLabel1.TabIndex = 19;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Yuta Hoshi (星 雄大) (hoshi5269@ihi-g.com)";
            this.linkLabel1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.linkLabel1.VisitedLinkColor = System.Drawing.Color.Green;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(161, 12);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(39, 20);
            this.label3.TabIndex = 18;
            this.label3.Text = "help";
            // 
            // Helpbtn
            // 
            this.Helpbtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(174)))), ((int)(((byte)(121)))));
            this.Helpbtn.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.Helpbtn.FlatAppearance.BorderSize = 0;
            this.Helpbtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Helpbtn.Font = new System.Drawing.Font("Yu Gothic UI", 12F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
                | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Helpbtn.ForeColor = System.Drawing.Color.White;
            this.Helpbtn.Location = new System.Drawing.Point(206, 6);
            this.Helpbtn.Name = "Helpbtn";
            this.Helpbtn.Size = new System.Drawing.Size(99, 33);
            this.Helpbtn.TabIndex = 17;
            this.Helpbtn.Text = "document.";
            this.Helpbtn.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.Helpbtn.UseVisualStyleBackColor = false;
            this.Helpbtn.Click += new System.EventHandler(this.Helpbtn_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(60, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 20);
            this.label2.TabIndex = 16;
            this.label2.Text = "Click Here for ";
            // 
            // pictureBox5
            // 
            this.pictureBox5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox5.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox5.BackgroundImage")));
            this.pictureBox5.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox5.Location = new System.Drawing.Point(694, -1);
            this.pictureBox5.Name = "pictureBox5";
            this.pictureBox5.Size = new System.Drawing.Size(77, 33);
            this.pictureBox5.TabIndex = 15;
            this.pictureBox5.TabStop = false;
            // 
            // LChart_Comparison
            // 
            this.LChart_Comparison.BackgroundImage = global::LChart_Comparison_Tool.Properties.Resources.Ticket_Maker;
            this.LChart_Comparison.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.LChart_Comparison.Controls.Add(this.btnParentChild);
            this.LChart_Comparison.Controls.Add(this.label16);
            this.LChart_Comparison.Controls.Add(this.progressBar1);
            this.LChart_Comparison.Controls.Add(this.textBox1);
            this.LChart_Comparison.Controls.Add(this.comboBox1);
            this.LChart_Comparison.Controls.Add(this.InputBrowse);
            this.LChart_Comparison.Controls.Add(this.InputPath);
            this.LChart_Comparison.Controls.Add(this.label13);
            this.LChart_Comparison.Controls.Add(this.BrowseBtn);
            this.LChart_Comparison.Controls.Add(this.label9);
            this.LChart_Comparison.Controls.Add(this.label10);
            this.LChart_Comparison.Controls.Add(this.pictureBox1);
            this.LChart_Comparison.Controls.Add(this.Form1_Update);
            this.LChart_Comparison.Controls.Add(this.label12);
            this.LChart_Comparison.Location = new System.Drawing.Point(4, 33);
            this.LChart_Comparison.Name = "LChart_Comparison";
            this.LChart_Comparison.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.LChart_Comparison.Size = new System.Drawing.Size(776, 320);
            this.LChart_Comparison.TabIndex = 6;
            this.LChart_Comparison.Text = "LChart Comparison   ";
            this.LChart_Comparison.UseVisualStyleBackColor = true;
            this.LChart_Comparison.Click += new System.EventHandler(this.Form1_Checklist_Click);
            // 
            // label16
            // 
            this.label16.Font = new System.Drawing.Font("Yu Gothic UI Semibold", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(70)))), ((int)(((byte)(172)))), ((int)(((byte)(103)))));
            this.label16.Location = new System.Drawing.Point(3, 252);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(503, 34);
            this.label16.TabIndex = 35;
            this.label16.Text = "Please Wait....";
            this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(-3, 289);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.progressBar1.Maximum = 150;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(781, 28);
            this.progressBar1.TabIndex = 34;
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(85)))), ((int)(((byte)(166)))), ((int)(((byte)(215)))));
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(194, 153);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(484, 50);
            this.textBox1.TabIndex = 33;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // comboBox1
            // 
            this.comboBox1.AllowDrop = true;
            this.comboBox1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.ForeColor = System.Drawing.SystemColors.InfoText;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.ItemHeight = 28;
            this.comboBox1.Location = new System.Drawing.Point(194, 46);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(388, 34);
            this.comboBox1.TabIndex = 32;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // InputBrowse
            // 
            this.InputBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.InputBrowse.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(151)))), ((int)(((byte)(148)))), ((int)(((byte)(26)))));
            this.InputBrowse.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.InputBrowse.FlatAppearance.BorderSize = 0;
            this.InputBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.InputBrowse.Font = new System.Drawing.Font("Arial Black", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.InputBrowse.ForeColor = System.Drawing.Color.Transparent;
            this.InputBrowse.Location = new System.Drawing.Point(682, 89);
            this.InputBrowse.Name = "InputBrowse";
            this.InputBrowse.Size = new System.Drawing.Size(52, 48);
            this.InputBrowse.TabIndex = 31;
            this.InputBrowse.Text = "......";
            this.InputBrowse.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.InputBrowse.UseVisualStyleBackColor = false;
            this.InputBrowse.Click += new System.EventHandler(this.FolderBrowse_Click);
            // 
            // InputPath
            // 
            this.InputPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.InputPath.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(85)))), ((int)(((byte)(166)))), ((int)(((byte)(215)))));
            this.InputPath.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.InputPath.Location = new System.Drawing.Point(194, 89);
            this.InputPath.Multiline = true;
            this.InputPath.Name = "InputPath";
            this.InputPath.Size = new System.Drawing.Size(484, 48);
            this.InputPath.TabIndex = 30;
            this.InputPath.TextChanged += new System.EventHandler(this.InputPath_TextChanged);
            // 
            // label13
            // 
            this.label13.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(202)))), ((int)(((byte)(182)))), ((int)(((byte)(226)))));
            this.label13.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label13.Font = new System.Drawing.Font("Yu Gothic UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(32, 89);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(162, 48);
            this.label13.TabIndex = 29;
            this.label13.Text = "Input Folder :";
            this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // BrowseBtn
            // 
            this.BrowseBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BrowseBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(151)))), ((int)(((byte)(148)))), ((int)(((byte)(26)))));
            this.BrowseBtn.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.BrowseBtn.FlatAppearance.BorderSize = 0;
            this.BrowseBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BrowseBtn.Font = new System.Drawing.Font("Arial Black", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BrowseBtn.ForeColor = System.Drawing.Color.Transparent;
            this.BrowseBtn.Location = new System.Drawing.Point(683, 153);
            this.BrowseBtn.Name = "BrowseBtn";
            this.BrowseBtn.Size = new System.Drawing.Size(51, 50);
            this.BrowseBtn.TabIndex = 28;
            this.BrowseBtn.Text = "......";
            this.BrowseBtn.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.BrowseBtn.UseVisualStyleBackColor = false;
            this.BrowseBtn.Click += new System.EventHandler(this.BrowseBtn_Click);
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(202)))), ((int)(((byte)(182)))), ((int)(((byte)(226)))));
            this.label9.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label9.Font = new System.Drawing.Font("Yu Gothic UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(33, 46);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(160, 28);
            this.label9.TabIndex = 24;
            this.label9.Text = "Type of Master :";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label10
            // 
            this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(6, 266);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(0, 20);
            this.label10.TabIndex = 23;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox1.Location = new System.Drawing.Point(700, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(77, 33);
            this.pictureBox1.TabIndex = 20;
            this.pictureBox1.TabStop = false;
            // 
            // Form1_Update
            // 
            this.Form1_Update.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.Form1_Update.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(23)))), ((int)(((byte)(136)))), ((int)(((byte)(111)))));
            this.Form1_Update.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.Form1_Update.FlatAppearance.BorderSize = 0;
            this.Form1_Update.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Form1_Update.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Form1_Update.ForeColor = System.Drawing.Color.White;
            this.Form1_Update.Location = new System.Drawing.Point(562, 232);
            this.Form1_Update.Name = "Form1_Update";
            this.Form1_Update.Size = new System.Drawing.Size(172, 39);
            this.Form1_Update.TabIndex = 16;
            this.Form1_Update.Text = "Update";
            this.Form1_Update.UseVisualStyleBackColor = false;
            this.Form1_Update.Click += new System.EventHandler(this.Form1_Update_Click);
            // 
            // label12
            // 
            this.label12.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(202)))), ((int)(((byte)(182)))), ((int)(((byte)(226)))));
            this.label12.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label12.Font = new System.Drawing.Font("Yu Gothic UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(32, 153);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(162, 50);
            this.label12.TabIndex = 26;
            this.label12.Text = "Cyient Portal File :";
            this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label12.Click += new System.EventHandler(this.label12_Click);
            // 
            // Tabs
            // 
            this.Tabs.Controls.Add(this.LChart_Comparison);
            this.Tabs.Controls.Add(this.Help);
            this.Tabs.Controls.Add(this.Version_Control);
            this.Tabs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Tabs.Font = new System.Drawing.Font("Yu Gothic UI Semibold", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Tabs.ItemSize = new System.Drawing.Size(125, 29);
            this.Tabs.Location = new System.Drawing.Point(0, 0);
            this.Tabs.Name = "Tabs";
            this.Tabs.SelectedIndex = 0;
            this.Tabs.Size = new System.Drawing.Size(784, 357);
            this.Tabs.TabIndex = 7;
            this.Tabs.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabControl1_DrawItem);
            // 
            // Version_Control
            // 
            this.Version_Control.BackColor = System.Drawing.Color.Teal;
            this.Version_Control.BackgroundImage = global::LChart_Comparison_Tool.Properties.Resources.LoginForm21;
            this.Version_Control.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Version_Control.Controls.Add(this.panel1);
            this.Version_Control.Controls.Add(this.pictureBox2);
            this.Version_Control.Location = new System.Drawing.Point(4, 33);
            this.Version_Control.Name = "Version_Control";
            this.Version_Control.Size = new System.Drawing.Size(776, 320);
            this.Version_Control.TabIndex = 5;
            this.Version_Control.Text = "   Version Control   ";
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.label14);
            this.panel1.Controls.Add(this.label15);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Location = new System.Drawing.Point(13, 40);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(750, 257);
            this.panel1.TabIndex = 18;
            // 
            // label14
            // 
            this.label14.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.label14.Location = new System.Drawing.Point(5, 411);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(728, 138);
            this.label14.TabIndex = 23;
            this.label14.Text = resources.GetString("label14.Text");
            // 
            // label15
            // 
            this.label15.BackColor = System.Drawing.Color.Transparent;
            this.label15.Font = new System.Drawing.Font("Yu Gothic UI", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.ForeColor = System.Drawing.Color.White;
            this.label15.Location = new System.Drawing.Point(0, 375);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(335, 33);
            this.label15.TabIndex = 22;
            this.label15.Text = "Gate1 Checklist Automation Tool V1.2";
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.label1.Location = new System.Drawing.Point(1, 213);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(728, 138);
            this.label1.TabIndex = 21;
            this.label1.Text = resources.GetString("label1.Text");
            // 
            // label11
            // 
            this.label11.BackColor = System.Drawing.Color.Transparent;
            this.label11.Font = new System.Drawing.Font("Yu Gothic UI", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.ForeColor = System.Drawing.Color.White;
            this.label11.Location = new System.Drawing.Point(-4, 177);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(335, 33);
            this.label11.TabIndex = 20;
            this.label11.Text = "Gate1 Checklist Automation Tool V1.1";
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.label6.Location = new System.Drawing.Point(-2, 35);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(731, 141);
            this.label6.TabIndex = 19;
            this.label6.Text = resources.GetString("label6.Text");
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.Font = new System.Drawing.Font("Yu Gothic UI", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(-3, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(352, 33);
            this.label5.TabIndex = 18;
            this.label5.Text = "Gate1 Checklist Automation Tool V1.0";
            // 
            // pictureBox2
            // 
            this.pictureBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox2.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox2.BackgroundImage")));
            this.pictureBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox2.Location = new System.Drawing.Point(698, 1);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(77, 33);
            this.pictureBox2.TabIndex = 21;
            this.pictureBox2.TabStop = false;
            // 
            // btnParentChild
            // 
            this.btnParentChild.Location = new System.Drawing.Point(302, 232);
            this.btnParentChild.Name = "btnParentChild";
            this.btnParentChild.Size = new System.Drawing.Size(215, 39);
            this.btnParentChild.TabIndex = 36;
            this.btnParentChild.Text = "Parent Child";
            this.btnParentChild.UseVisualStyleBackColor = true;
            this.btnParentChild.Click += new System.EventHandler(this.btnParentChild_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 357);
            this.Controls.Add(this.Tabs);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "LChart_Comparison_Tool_V1.0";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Resize += new System.EventHandler(this.Form1_Resize);
            this.Help.ResumeLayout(false);
            this.Help.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).EndInit();
            this.LChart_Comparison.ResumeLayout(false);
            this.LChart_Comparison.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.Tabs.ResumeLayout(false);
            this.Version_Control.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);

        }

       


        #endregion

        private System.Windows.Forms.TabPage Help;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.LinkLabel linkLabel3;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button Helpbtn;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox pictureBox5;
        private System.Windows.Forms.TabPage LChart_Comparison;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button Form1_Update;
        private System.Windows.Forms.TabControl Tabs;
        private System.Windows.Forms.TabPage Version_Control;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button BrowseBtn;
        private System.Windows.Forms.Button InputBrowse;
        private System.Windows.Forms.TextBox InputPath;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.TextBox textBox1;
        private ProgressBar progressBar1;
        private Label label16;
        private Label label13;
        private Button btnParentChild;
    }
}