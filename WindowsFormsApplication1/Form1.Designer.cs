namespace WindowsFormsApplication1
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
      this.components = new System.ComponentModel.Container();
      this.panel1 = new System.Windows.Forms.Panel();
      this.button7 = new System.Windows.Forms.Button();
      this.button6 = new System.Windows.Forms.Button();
      this.label4 = new System.Windows.Forms.Label();
      this.label3 = new System.Windows.Forms.Label();
      this.buttonPopulateSheetsWithVaulters = new System.Windows.Forms.Button();
      this.label2 = new System.Windows.Forms.Label();
      this.buttonCreateResultSheets = new System.Windows.Forms.Button();
      this.btnReadResultsFromInbox = new System.Windows.Forms.Button();
      this.label1 = new System.Windows.Forms.Label();
      this.buttonFakeResults = new System.Windows.Forms.Button();
      this.panel3 = new System.Windows.Forms.Panel();
      this.buttonClear = new System.Windows.Forms.Button();
      this.textBox1 = new System.Windows.Forms.TextBox();
      this.progressLabel = new System.Windows.Forms.Label();
      this.progressBar1 = new System.Windows.Forms.ProgressBar();
      this.backgroundWorkerFakeResults = new System.ComponentModel.BackgroundWorker();
      this.tabControl1 = new System.Windows.Forms.TabControl();
      this.tabPage1 = new System.Windows.Forms.TabPage();
      this.dataGridView1 = new System.Windows.Forms.DataGridView();
      this.tabPage2 = new System.Windows.Forms.TabPage();
      this.dataGridView2 = new System.Windows.Forms.DataGridView();
      this.tabPage3 = new System.Windows.Forms.TabPage();
      this.dataGridView3 = new System.Windows.Forms.DataGridView();
      this.button4 = new System.Windows.Forms.Button();
      this.panel2 = new System.Windows.Forms.Panel();
      this.checkBoxJudge = new System.Windows.Forms.CheckBox();
      this.textBoxProcessInterval = new System.Windows.Forms.TextBox();
      this.checkBoxProcessTimer = new System.Windows.Forms.CheckBox();
      this.button5 = new System.Windows.Forms.Button();
      this.button2 = new System.Windows.Forms.Button();
      this.panel4 = new System.Windows.Forms.Panel();
      this.createPdfsCheckBox = new System.Windows.Forms.CheckBox();
      this.button3 = new System.Windows.Forms.Button();
      this.button1 = new System.Windows.Forms.Button();
      this.label5 = new System.Windows.Forms.Label();
      this.comboBox1 = new System.Windows.Forms.ComboBox();
      this.checkBox1 = new System.Windows.Forms.CheckBox();
      this.backgroundWorkerCreateClassResultsSheets = new System.ComponentModel.BackgroundWorker();
      this.backgroundWorkerPopulateSheetsWithVaulters = new System.ComponentModel.BackgroundWorker();
      this.backgroundWorkerReadResultsFromInbox = new System.ComponentModel.BackgroundWorker();
      this.backgroundWorker5 = new System.ComponentModel.BackgroundWorker();
      this.backgroundWorkerSortResults = new System.ComponentModel.BackgroundWorker();
      this.printDialog1 = new System.Windows.Forms.PrintDialog();
      this.backgroundWorkerPrintResults = new System.ComponentModel.BackgroundWorker();
      this.backgroundWorkerPublish = new System.ComponentModel.BackgroundWorker();
      this.processResultsTimer = new System.Windows.Forms.Timer(this.components);
      this.backgroundWorkerFullAutoProcess = new System.ComponentModel.BackgroundWorker();
      this.backgroundWorkerJudgeTables = new System.ComponentModel.BackgroundWorker();
      this.judgeTimer = new System.Windows.Forms.Timer(this.components);
      this.panel1.SuspendLayout();
      this.panel3.SuspendLayout();
      this.tabControl1.SuspendLayout();
      this.tabPage1.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
      this.tabPage2.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
      this.tabPage3.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
      this.panel2.SuspendLayout();
      this.panel4.SuspendLayout();
      this.SuspendLayout();
      // 
      // panel1
      // 
      this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.panel1.Controls.Add(this.button7);
      this.panel1.Controls.Add(this.button6);
      this.panel1.Controls.Add(this.label4);
      this.panel1.Controls.Add(this.label3);
      this.panel1.Controls.Add(this.buttonPopulateSheetsWithVaulters);
      this.panel1.Controls.Add(this.label2);
      this.panel1.Controls.Add(this.buttonCreateResultSheets);
      this.panel1.Controls.Add(this.btnReadResultsFromInbox);
      this.panel1.Controls.Add(this.label1);
      this.panel1.Controls.Add(this.buttonFakeResults);
      this.panel1.Controls.Add(this.panel3);
      this.panel1.Location = new System.Drawing.Point(30, 527);
      this.panel1.Name = "panel1";
      this.panel1.Size = new System.Drawing.Size(1458, 248);
      this.panel1.TabIndex = 0;
      // 
      // button7
      // 
      this.button7.Location = new System.Drawing.Point(202, 143);
      this.button7.Name = "button7";
      this.button7.Size = new System.Drawing.Size(75, 23);
      this.button7.TabIndex = 15;
      this.button7.Text = "häst ID";
      this.button7.UseVisualStyleBackColor = true;
      this.button7.Visible = false;
      this.button7.Click += new System.EventHandler(this.button7_Click);
      // 
      // button6
      // 
      this.button6.Location = new System.Drawing.Point(75, 178);
      this.button6.Name = "button6";
      this.button6.Size = new System.Drawing.Size(216, 42);
      this.button6.TabIndex = 14;
      this.button6.Text = "Omvänd startordning";
      this.button6.UseVisualStyleBackColor = true;
      this.button6.Visible = false;
      this.button6.Click += new System.EventHandler(this.button6_Click);
      // 
      // label4
      // 
      this.label4.AutoSize = true;
      this.label4.Location = new System.Drawing.Point(31, 118);
      this.label4.Name = "label4";
      this.label4.Size = new System.Drawing.Size(38, 13);
      this.label4.TabIndex = 13;
      this.label4.Text = "Step 5";
      // 
      // label3
      // 
      this.label3.AutoSize = true;
      this.label3.Location = new System.Drawing.Point(31, 56);
      this.label3.Name = "label3";
      this.label3.Size = new System.Drawing.Size(38, 13);
      this.label3.TabIndex = 12;
      this.label3.Text = "Step 3";
      // 
      // buttonPopulateSheetsWithVaulters
      // 
      this.buttonPopulateSheetsWithVaulters.Location = new System.Drawing.Point(75, 51);
      this.buttonPopulateSheetsWithVaulters.Name = "buttonPopulateSheetsWithVaulters";
      this.buttonPopulateSheetsWithVaulters.Size = new System.Drawing.Size(217, 23);
      this.buttonPopulateSheetsWithVaulters.TabIndex = 11;
      this.buttonPopulateSheetsWithVaulters.Text = "Populate Results sheets with vaulters";
      this.buttonPopulateSheetsWithVaulters.UseVisualStyleBackColor = true;
      this.buttonPopulateSheetsWithVaulters.Click += new System.EventHandler(this.buttonPopulateSheetsWithVaulters_Click);
      // 
      // label2
      // 
      this.label2.AutoSize = true;
      this.label2.Location = new System.Drawing.Point(31, 25);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(38, 13);
      this.label2.TabIndex = 8;
      this.label2.Text = "Step 2";
      // 
      // buttonCreateResultSheets
      // 
      this.buttonCreateResultSheets.Location = new System.Drawing.Point(75, 20);
      this.buttonCreateResultSheets.Name = "buttonCreateResultSheets";
      this.buttonCreateResultSheets.Size = new System.Drawing.Size(217, 23);
      this.buttonCreateResultSheets.TabIndex = 7;
      this.buttonCreateResultSheets.Text = "Create Base Result Sheets for all classes";
      this.buttonCreateResultSheets.UseVisualStyleBackColor = true;
      this.buttonCreateResultSheets.Click += new System.EventHandler(this.buttonCreateResultSheets_Click);
      // 
      // btnReadResultsFromInbox
      // 
      this.btnReadResultsFromInbox.Location = new System.Drawing.Point(75, 113);
      this.btnReadResultsFromInbox.Name = "btnReadResultsFromInbox";
      this.btnReadResultsFromInbox.Size = new System.Drawing.Size(217, 23);
      this.btnReadResultsFromInbox.TabIndex = 6;
      this.btnReadResultsFromInbox.Text = "Read results from Inbox and sort";
      this.btnReadResultsFromInbox.UseVisualStyleBackColor = true;
      this.btnReadResultsFromInbox.Click += new System.EventHandler(this.btnReadResultsFromInbox_Click);
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Location = new System.Drawing.Point(31, 87);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(38, 13);
      this.label1.TabIndex = 5;
      this.label1.Text = "Step 4";
      // 
      // buttonFakeResults
      // 
      this.buttonFakeResults.Location = new System.Drawing.Point(75, 82);
      this.buttonFakeResults.Name = "buttonFakeResults";
      this.buttonFakeResults.Size = new System.Drawing.Size(217, 23);
      this.buttonFakeResults.TabIndex = 4;
      this.buttonFakeResults.Text = "Fake Results";
      this.buttonFakeResults.UseVisualStyleBackColor = true;
      this.buttonFakeResults.Click += new System.EventHandler(this.buttonFakeResults_Click);
      // 
      // panel3
      // 
      this.panel3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.panel3.Controls.Add(this.buttonClear);
      this.panel3.Controls.Add(this.textBox1);
      this.panel3.Controls.Add(this.progressLabel);
      this.panel3.Controls.Add(this.progressBar1);
      this.panel3.Location = new System.Drawing.Point(327, 20);
      this.panel3.Name = "panel3";
      this.panel3.Size = new System.Drawing.Size(1106, 215);
      this.panel3.TabIndex = 10;
      // 
      // buttonClear
      // 
      this.buttonClear.Anchor = System.Windows.Forms.AnchorStyles.Right;
      this.buttonClear.Location = new System.Drawing.Point(1024, 59);
      this.buttonClear.Name = "buttonClear";
      this.buttonClear.Size = new System.Drawing.Size(51, 23);
      this.buttonClear.TabIndex = 11;
      this.buttonClear.Text = "Clear";
      this.buttonClear.UseVisualStyleBackColor = true;
      this.buttonClear.Click += new System.EventHandler(this.buttonClear_Click);
      // 
      // textBox1
      // 
      this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.textBox1.Location = new System.Drawing.Point(14, 59);
      this.textBox1.Multiline = true;
      this.textBox1.Name = "textBox1";
      this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
      this.textBox1.Size = new System.Drawing.Size(993, 153);
      this.textBox1.TabIndex = 10;
      // 
      // progressLabel
      // 
      this.progressLabel.AutoSize = true;
      this.progressLabel.Location = new System.Drawing.Point(11, 10);
      this.progressLabel.Name = "progressLabel";
      this.progressLabel.Size = new System.Drawing.Size(48, 13);
      this.progressLabel.TabIndex = 9;
      this.progressLabel.Text = "Progress";
      this.progressLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
      // 
      // progressBar1
      // 
      this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.progressBar1.Location = new System.Drawing.Point(14, 26);
      this.progressBar1.Name = "progressBar1";
      this.progressBar1.Size = new System.Drawing.Size(993, 23);
      this.progressBar1.TabIndex = 3;
      // 
      // backgroundWorkerFakeResults
      // 
      this.backgroundWorkerFakeResults.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
      // 
      // tabControl1
      // 
      this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.tabControl1.Controls.Add(this.tabPage1);
      this.tabControl1.Controls.Add(this.tabPage2);
      this.tabControl1.Controls.Add(this.tabPage3);
      this.tabControl1.Location = new System.Drawing.Point(17, 104);
      this.tabControl1.Name = "tabControl1";
      this.tabControl1.SelectedIndex = 0;
      this.tabControl1.Size = new System.Drawing.Size(1425, 378);
      this.tabControl1.TabIndex = 1;
      // 
      // tabPage1
      // 
      this.tabPage1.Controls.Add(this.dataGridView1);
      this.tabPage1.Location = new System.Drawing.Point(4, 22);
      this.tabPage1.Name = "tabPage1";
      this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
      this.tabPage1.Size = new System.Drawing.Size(1417, 352);
      this.tabPage1.TabIndex = 0;
      this.tabPage1.Text = "tabPage1";
      this.tabPage1.UseVisualStyleBackColor = true;
      // 
      // dataGridView1
      // 
      this.dataGridView1.AllowUserToAddRows = false;
      this.dataGridView1.AllowUserToDeleteRows = false;
      this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
      this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dataGridView1.Location = new System.Drawing.Point(6, 6);
      this.dataGridView1.Name = "dataGridView1";
      this.dataGridView1.ReadOnly = true;
      this.dataGridView1.RowHeadersWidth = 51;
      this.dataGridView1.Size = new System.Drawing.Size(1399, 372);
      this.dataGridView1.TabIndex = 0;
      // 
      // tabPage2
      // 
      this.tabPage2.Controls.Add(this.dataGridView2);
      this.tabPage2.Location = new System.Drawing.Point(4, 22);
      this.tabPage2.Name = "tabPage2";
      this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
      this.tabPage2.Size = new System.Drawing.Size(1417, 352);
      this.tabPage2.TabIndex = 1;
      this.tabPage2.Text = "tabPage2";
      this.tabPage2.UseVisualStyleBackColor = true;
      // 
      // dataGridView2
      // 
      this.dataGridView2.AllowUserToAddRows = false;
      this.dataGridView2.AllowUserToDeleteRows = false;
      this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
      this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dataGridView2.Location = new System.Drawing.Point(7, 6);
      this.dataGridView2.Name = "dataGridView2";
      this.dataGridView2.ReadOnly = true;
      this.dataGridView2.RowHeadersWidth = 51;
      this.dataGridView2.Size = new System.Drawing.Size(1404, 372);
      this.dataGridView2.TabIndex = 0;
      // 
      // tabPage3
      // 
      this.tabPage3.Controls.Add(this.dataGridView3);
      this.tabPage3.Location = new System.Drawing.Point(4, 22);
      this.tabPage3.Name = "tabPage3";
      this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
      this.tabPage3.Size = new System.Drawing.Size(1417, 352);
      this.tabPage3.TabIndex = 2;
      this.tabPage3.Text = "tabPage3";
      this.tabPage3.UseVisualStyleBackColor = true;
      // 
      // dataGridView3
      // 
      this.dataGridView3.AllowUserToAddRows = false;
      this.dataGridView3.AllowUserToDeleteRows = false;
      this.dataGridView3.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
      this.dataGridView3.BackgroundColor = System.Drawing.Color.Azure;
      this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dataGridView3.Location = new System.Drawing.Point(7, 6);
      this.dataGridView3.Name = "dataGridView3";
      this.dataGridView3.ReadOnly = true;
      this.dataGridView3.RowHeadersWidth = 51;
      this.dataGridView3.RowTemplate.Height = 16;
      this.dataGridView3.Size = new System.Drawing.Size(1404, 372);
      this.dataGridView3.TabIndex = 1;
      this.dataGridView3.DataSourceChanged += new System.EventHandler(this.dataGridView3_DataSourceChanged);
      this.dataGridView3.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridView3_CellFormatting);
      this.dataGridView3.RowPrePaint += new System.Windows.Forms.DataGridViewRowPrePaintEventHandler(this.dataGridView3_RowPrePaint);
      // 
      // button4
      // 
      this.button4.Location = new System.Drawing.Point(17, 16);
      this.button4.Name = "button4";
      this.button4.Size = new System.Drawing.Size(395, 31);
      this.button4.TabIndex = 7;
      this.button4.Text = "Step 1 - Read Classes and Vaulters from Startlist";
      this.button4.UseVisualStyleBackColor = true;
      this.button4.Click += new System.EventHandler(this.button4_Click);
      // 
      // panel2
      // 
      this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.panel2.Controls.Add(this.checkBoxJudge);
      this.panel2.Controls.Add(this.textBoxProcessInterval);
      this.panel2.Controls.Add(this.checkBoxProcessTimer);
      this.panel2.Controls.Add(this.button5);
      this.panel2.Controls.Add(this.button2);
      this.panel2.Controls.Add(this.button4);
      this.panel2.Controls.Add(this.panel4);
      this.panel2.Controls.Add(this.tabControl1);
      this.panel2.Location = new System.Drawing.Point(30, 26);
      this.panel2.Name = "panel2";
      this.panel2.Size = new System.Drawing.Size(1458, 495);
      this.panel2.TabIndex = 8;
      // 
      // checkBoxJudge
      // 
      this.checkBoxJudge.AutoSize = true;
      this.checkBoxJudge.Enabled = false;
      this.checkBoxJudge.Location = new System.Drawing.Point(121, 54);
      this.checkBoxJudge.Name = "checkBoxJudge";
      this.checkBoxJudge.Size = new System.Drawing.Size(117, 17);
      this.checkBoxJudge.TabIndex = 16;
      this.checkBoxJudge.Text = "Judge Table Points";
      this.checkBoxJudge.UseVisualStyleBackColor = true;
      this.checkBoxJudge.Visible = false;
      this.checkBoxJudge.CheckedChanged += new System.EventHandler(this.checkBoxJudge_CheckedChanged);
      // 
      // textBoxProcessInterval
      // 
      this.textBoxProcessInterval.Enabled = false;
      this.textBoxProcessInterval.Location = new System.Drawing.Point(310, 76);
      this.textBoxProcessInterval.Name = "textBoxProcessInterval";
      this.textBoxProcessInterval.Size = new System.Drawing.Size(100, 20);
      this.textBoxProcessInterval.TabIndex = 15;
      this.textBoxProcessInterval.Text = "600";
      this.textBoxProcessInterval.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
      this.textBoxProcessInterval.Visible = false;
      this.textBoxProcessInterval.TextChanged += new System.EventHandler(this.textBoxProcessInterval_TextChanged);
      // 
      // checkBoxProcessTimer
      // 
      this.checkBoxProcessTimer.AutoSize = true;
      this.checkBoxProcessTimer.Enabled = false;
      this.checkBoxProcessTimer.Location = new System.Drawing.Point(310, 53);
      this.checkBoxProcessTimer.Name = "checkBoxProcessTimer";
      this.checkBoxProcessTimer.Size = new System.Drawing.Size(102, 17);
      this.checkBoxProcessTimer.TabIndex = 9;
      this.checkBoxProcessTimer.Text = "Auto processing";
      this.checkBoxProcessTimer.UseVisualStyleBackColor = true;
      this.checkBoxProcessTimer.Visible = false;
      this.checkBoxProcessTimer.CheckedChanged += new System.EventHandler(this.checkBoxProcessTimer_CheckedChanged);
      // 
      // button5
      // 
      this.button5.Anchor = System.Windows.Forms.AnchorStyles.Right;
      this.button5.Location = new System.Drawing.Point(1008, 43);
      this.button5.Name = "button5";
      this.button5.Size = new System.Drawing.Size(202, 23);
      this.button5.TabIndex = 11;
      this.button5.Text = "Export results for all classes";
      this.button5.UseVisualStyleBackColor = true;
      this.button5.Click += new System.EventHandler(this.button5_Click);
      // 
      // button2
      // 
      this.button2.Anchor = System.Windows.Forms.AnchorStyles.Right;
      this.button2.Location = new System.Drawing.Point(1008, 18);
      this.button2.Name = "button2";
      this.button2.Size = new System.Drawing.Size(202, 23);
      this.button2.TabIndex = 10;
      this.button2.Text = "Export result for selected  class";
      this.button2.UseVisualStyleBackColor = true;
      this.button2.Click += new System.EventHandler(this.button2_Click_1);
      // 
      // panel4
      // 
      this.panel4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.panel4.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
      this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.panel4.Controls.Add(this.createPdfsCheckBox);
      this.panel4.Controls.Add(this.button3);
      this.panel4.Controls.Add(this.button1);
      this.panel4.Controls.Add(this.label5);
      this.panel4.Controls.Add(this.comboBox1);
      this.panel4.Controls.Add(this.checkBox1);
      this.panel4.Location = new System.Drawing.Point(440, 6);
      this.panel4.Name = "panel4";
      this.panel4.Size = new System.Drawing.Size(992, 92);
      this.panel4.TabIndex = 14;
      // 
      // createPdfsCheckBox
      // 
      this.createPdfsCheckBox.AutoSize = true;
      this.createPdfsCheckBox.Location = new System.Drawing.Point(373, 43);
      this.createPdfsCheckBox.Name = "createPdfsCheckBox";
      this.createPdfsCheckBox.Size = new System.Drawing.Size(150, 17);
      this.createPdfsCheckBox.TabIndex = 16;
      this.createPdfsCheckBox.Text = "Create PDFs during export";
      this.createPdfsCheckBox.UseVisualStyleBackColor = true;
      this.createPdfsCheckBox.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
      // 
      // button3
      // 
      this.button3.Anchor = System.Windows.Forms.AnchorStyles.Right;
      this.button3.Location = new System.Drawing.Point(775, 39);
      this.button3.Name = "button3";
      this.button3.Size = new System.Drawing.Size(186, 22);
      this.button3.TabIndex = 15;
      this.button3.Text = "Calculate horse points";
      this.button3.UseVisualStyleBackColor = true;
      this.button3.Click += new System.EventHandler(this.button3_Click);
      // 
      // button1
      // 
      this.button1.Anchor = System.Windows.Forms.AnchorStyles.Right;
      this.button1.Location = new System.Drawing.Point(775, 11);
      this.button1.Name = "button1";
      this.button1.Size = new System.Drawing.Size(186, 26);
      this.button1.TabIndex = 13;
      this.button1.Text = "Publish";
      this.button1.UseVisualStyleBackColor = true;
      this.button1.Click += new System.EventHandler(this.button1_Click);
      // 
      // label5
      // 
      this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
      this.label5.AutoSize = true;
      this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.label5.Location = new System.Drawing.Point(23, 10);
      this.label5.Name = "label5";
      this.label5.Size = new System.Drawing.Size(120, 13);
      this.label5.TabIndex = 13;
      this.label5.Text = "Results and printing";
      // 
      // comboBox1
      // 
      this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
      this.comboBox1.FormattingEnabled = true;
      this.comboBox1.Location = new System.Drawing.Point(26, 39);
      this.comboBox1.Name = "comboBox1";
      this.comboBox1.Size = new System.Drawing.Size(325, 21);
      this.comboBox1.TabIndex = 9;
      this.comboBox1.Text = "Select Class";
      this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
      this.comboBox1.SelectedValueChanged += new System.EventHandler(this.comboBox1_SelectedValueChanged);
      // 
      // checkBox1
      // 
      this.checkBox1.Anchor = System.Windows.Forms.AnchorStyles.Right;
      this.checkBox1.AutoSize = true;
      this.checkBox1.Checked = true;
      this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
      this.checkBox1.Location = new System.Drawing.Point(373, 15);
      this.checkBox1.Name = "checkBox1";
      this.checkBox1.Size = new System.Drawing.Size(166, 17);
      this.checkBox1.TabIndex = 12;
      this.checkBox1.Text = "Add \"Preliminiary result\" label ";
      this.checkBox1.UseVisualStyleBackColor = true;
      this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
      // 
      // backgroundWorkerCreateClassResultsSheets
      // 
      this.backgroundWorkerCreateClassResultsSheets.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerCreateClassResultsSheets_DoWork);
      this.backgroundWorkerCreateClassResultsSheets.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerCreateClassResultsSheets_RunWorkerCompleted);
      // 
      // backgroundWorkerPopulateSheetsWithVaulters
      // 
      this.backgroundWorkerPopulateSheetsWithVaulters.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerPopulateSheetsWithVaulters_DoWork);
      this.backgroundWorkerPopulateSheetsWithVaulters.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerPopulateSheetsWithVaulters_RunWorkerCompleted);
      // 
      // backgroundWorkerReadResultsFromInbox
      // 
      this.backgroundWorkerReadResultsFromInbox.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerReadResultsFromInbox_DoWork);
      this.backgroundWorkerReadResultsFromInbox.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerReadResultsFromInbox_RunWorkerCompleted);
      // 
      // backgroundWorker5
      // 
      this.backgroundWorker5.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker5_DoWork);
      this.backgroundWorker5.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker5_ProgressChanged);
      // 
      // backgroundWorkerSortResults
      // 
      this.backgroundWorkerSortResults.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerSortResults_DoWork);
      this.backgroundWorkerSortResults.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerSortResults_RunWorkerCompleted);
      // 
      // printDialog1
      // 
      this.printDialog1.UseEXDialog = true;
      // 
      // backgroundWorkerPrintResults
      // 
      this.backgroundWorkerPrintResults.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerPrintResults_DoWork);
      this.backgroundWorkerPrintResults.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerPrintResults_RunWorkerCompleted);
      // 
      // backgroundWorkerPublish
      // 
      this.backgroundWorkerPublish.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerPublish_DoWork);
      this.backgroundWorkerPublish.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerPublish_RunWorkerCompleted);
      // 
      // processResultsTimer
      // 
      this.processResultsTimer.Interval = 30000;
      this.processResultsTimer.Tick += new System.EventHandler(this.processResultsTimer_Tick);
      // 
      // backgroundWorkerFullAutoProcess
      // 
      this.backgroundWorkerFullAutoProcess.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerFullAutoProcess_DoWork);
      this.backgroundWorkerFullAutoProcess.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerFullAutoProcess_RunWorkerCompleted);
      // 
      // backgroundWorkerJudgeTables
      // 
      this.backgroundWorkerJudgeTables.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerJudgeTables_DoWork);
      this.backgroundWorkerJudgeTables.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerJudgeTables_RunWorkerCompleted);
      // 
      // judgeTimer
      // 
      this.judgeTimer.Tick += new System.EventHandler(this.judgeTimer_Tick);
      // 
      // Form1
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(1500, 818);
      this.Controls.Add(this.panel2);
      this.Controls.Add(this.panel1);
      this.Name = "Form1";
      this.Text = "Form1";
      this.panel1.ResumeLayout(false);
      this.panel1.PerformLayout();
      this.panel3.ResumeLayout(false);
      this.panel3.PerformLayout();
      this.tabControl1.ResumeLayout(false);
      this.tabPage1.ResumeLayout(false);
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
      this.tabPage2.ResumeLayout(false);
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
      this.tabPage3.ResumeLayout(false);
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
      this.panel2.ResumeLayout(false);
      this.panel2.PerformLayout();
      this.panel4.ResumeLayout(false);
      this.panel4.PerformLayout();
      this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button buttonFakeResults;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnReadResultsFromInbox;
        private System.ComponentModel.BackgroundWorker backgroundWorkerFakeResults;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonCreateResultSheets;
        private System.Windows.Forms.Label progressLabel;
        private System.Windows.Forms.Panel panel3;
        private System.ComponentModel.BackgroundWorker backgroundWorkerCreateClassResultsSheets;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button buttonPopulateSheetsWithVaulters;
        private System.Windows.Forms.Label label4;
        private System.ComponentModel.BackgroundWorker backgroundWorkerReadResultsFromInbox;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.DataGridView dataGridView3;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button button2;
        private System.ComponentModel.BackgroundWorker backgroundWorker5;
        private System.ComponentModel.BackgroundWorker backgroundWorkerPopulateSheetsWithVaulters;

        private System.Windows.Forms.Button button5;
        private System.ComponentModel.BackgroundWorker backgroundWorkerSortResults;
        private System.Windows.Forms.PrintDialog printDialog1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button buttonClear;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Button button1;
    private System.Windows.Forms.Panel panel4;
    private System.Windows.Forms.Label label5;
    private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button6;
    private System.Windows.Forms.CheckBox createPdfsCheckBox;
    private System.ComponentModel.BackgroundWorker backgroundWorkerPrintResults;
    private System.ComponentModel.BackgroundWorker backgroundWorkerPublish;
    private System.Windows.Forms.Timer processResultsTimer;
    private System.Windows.Forms.CheckBox checkBoxProcessTimer;
    private System.Windows.Forms.TextBox textBoxProcessInterval;
    private System.ComponentModel.BackgroundWorker backgroundWorkerFullAutoProcess;
    private System.ComponentModel.BackgroundWorker backgroundWorkerJudgeTables;
    private System.Windows.Forms.CheckBox checkBoxJudge;
    private System.Windows.Forms.Timer judgeTimer;
    private System.Windows.Forms.Button button7;
  }
}

