namespace BikeSharingSystem
{
    partial class MainForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series5 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series6 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series7 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series8 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label4 = new System.Windows.Forms.Label();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.labelMax = new System.Windows.Forms.Label();
            this.labelBO = new System.Windows.Forms.Label();
            this.chart = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.butpga = new System.Windows.Forms.Button();
            this.ppg = new System.Windows.Forms.PropertyGrid();
            this.butreset = new System.Windows.Forms.Button();
            this.butrunend = new System.Windows.Forms.Button();
            this.tabPage = new System.Windows.Forms.TabPage();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.tabControl3 = new System.Windows.Forms.TabControl();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txbtimehorizon = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txbcapacity = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxPnDTime = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txbspeed = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.labelIteration = new System.Windows.Forms.Label();
            this.checkBox_times = new System.Windows.Forms.CheckBox();
            this.txbCPUTime = new System.Windows.Forms.TextBox();
            this.labelfilename = new System.Windows.Forms.Label();
            this.labeltitle = new System.Windows.Forms.Label();
            this.checkBoxGreedy = new System.Windows.Forms.CheckBox();
            this.tabPage6 = new System.Windows.Forms.TabPage();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton2 = new System.Windows.Forms.ToolStripButton();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.checkBox = new System.Windows.Forms.CheckBox();
            this.tabControl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart)).BeginInit();
            this.tabPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.tabControl3.SuspendLayout();
            this.tabPage5.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl
            // 
            this.tabControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl.Controls.Add(this.tabPage1);
            this.tabControl.Controls.Add(this.tabPage);
            this.tabControl.Controls.Add(this.tabPage2);
            this.tabControl.Location = new System.Drawing.Point(4, 1);
            this.tabControl.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(2341, 977);
            this.tabControl.TabIndex = 2;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.dataGridView);
            this.tabPage1.Controls.Add(this.labelMax);
            this.tabPage1.Controls.Add(this.labelBO);
            this.tabPage1.Controls.Add(this.chart);
            this.tabPage1.Controls.Add(this.butpga);
            this.tabPage1.Controls.Add(this.ppg);
            this.tabPage1.Controls.Add(this.butreset);
            this.tabPage1.Controls.Add(this.butrunend);
            this.tabPage1.Location = new System.Drawing.Point(4, 34);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage1.Size = new System.Drawing.Size(2333, 939);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "GA Optimization";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(340, 465);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(340, 25);
            this.label4.TabIndex = 14;
            this.label4.Text = "unsatisfied amount  of each station";
            // 
            // dataGridView
            // 
            this.dataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Location = new System.Drawing.Point(345, 493);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowTemplate.Height = 27;
            this.dataGridView.Size = new System.Drawing.Size(1113, 332);
            this.dataGridView.TabIndex = 13;
            // 
            // labelMax
            // 
            this.labelMax.AutoSize = true;
            this.labelMax.Location = new System.Drawing.Point(926, 27);
            this.labelMax.Name = "labelMax";
            this.labelMax.Size = new System.Drawing.Size(62, 25);
            this.labelMax.TabIndex = 12;
            this.labelMax.Text = "Max :";
            // 
            // labelBO
            // 
            this.labelBO.AutoSize = true;
            this.labelBO.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labelBO.ForeColor = System.Drawing.Color.Navy;
            this.labelBO.Location = new System.Drawing.Point(355, 27);
            this.labelBO.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelBO.Name = "labelBO";
            this.labelBO.Size = new System.Drawing.Size(117, 25);
            this.labelBO.TabIndex = 8;
            this.labelBO.Text = "BestObject";
            // 
            // chart
            // 
            this.chart.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            chartArea2.AxisX.Title = "iteration";
            chartArea2.AxisX2.Title = "objective";
            chartArea2.AxisY.Title = "objective";
            chartArea2.AxisY2.Title = "objective";
            chartArea2.BackColor = System.Drawing.Color.White;
            chartArea2.Name = "ChartArea1";
            this.chart.ChartAreas.Add(chartArea2);
            legend2.Docking = System.Windows.Forms.DataVisualization.Charting.Docking.Top;
            legend2.Name = "Legend1";
            this.chart.Legends.Add(legend2);
            this.chart.Location = new System.Drawing.Point(297, 57);
            this.chart.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chart.Name = "chart";
            series5.BorderColor = System.Drawing.Color.Black;
            series5.BorderWidth = 5;
            series5.ChartArea = "ChartArea1";
            series5.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series5.Color = System.Drawing.Color.Lime;
            series5.Legend = "Legend1";
            series5.MarkerBorderWidth = 5;
            series5.Name = "iteration average";
            series6.BorderWidth = 5;
            series6.ChartArea = "ChartArea1";
            series6.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series6.Color = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            series6.Legend = "Legend1";
            series6.Name = "iteration best";
            series7.BorderWidth = 5;
            series7.ChartArea = "ChartArea1";
            series7.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series7.Color = System.Drawing.Color.Red;
            series7.Legend = "Legend1";
            series7.Name = "sofar the best";
            series8.BorderDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Dash;
            series8.BorderWidth = 5;
            series8.ChartArea = "ChartArea1";
            series8.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series8.Color = System.Drawing.Color.DarkOrchid;
            series8.Legend = "Legend1";
            series8.Name = "not do anything unsatisfied amount";
            this.chart.Series.Add(series5);
            this.chart.Series.Add(series6);
            this.chart.Series.Add(series7);
            this.chart.Series.Add(series8);
            this.chart.Size = new System.Drawing.Size(1161, 441);
            this.chart.TabIndex = 8;
            this.chart.Text = "chart";
            // 
            // butpga
            // 
            this.butpga.Font = new System.Drawing.Font("新細明體", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.butpga.Location = new System.Drawing.Point(31, 39);
            this.butpga.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.butpga.Name = "butpga";
            this.butpga.Size = new System.Drawing.Size(249, 78);
            this.butpga.TabIndex = 7;
            this.butpga.Text = "create permutation GA";
            this.butpga.UseVisualStyleBackColor = true;
            this.butpga.Click += new System.EventHandler(this.createPGA);
            // 
            // ppg
            // 
            this.ppg.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.ppg.CategoryForeColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.ppg.Location = new System.Drawing.Point(31, 299);
            this.ppg.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.ppg.Name = "ppg";
            this.ppg.Size = new System.Drawing.Size(249, 526);
            this.ppg.TabIndex = 7;
            // 
            // butreset
            // 
            this.butreset.Font = new System.Drawing.Font("新細明體", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.butreset.Location = new System.Drawing.Point(31, 127);
            this.butreset.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.butreset.Name = "butreset";
            this.butreset.Size = new System.Drawing.Size(249, 72);
            this.butreset.TabIndex = 10;
            this.butreset.Text = "reset";
            this.butreset.UseVisualStyleBackColor = true;
            this.butreset.Click += new System.EventHandler(this.reset);
            // 
            // butrunend
            // 
            this.butrunend.Font = new System.Drawing.Font("新細明體", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.butrunend.Location = new System.Drawing.Point(31, 209);
            this.butrunend.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.butrunend.Name = "butrunend";
            this.butrunend.Size = new System.Drawing.Size(249, 80);
            this.butrunend.TabIndex = 9;
            this.butrunend.Text = "run";
            this.butrunend.UseVisualStyleBackColor = true;
            this.butrunend.Click += new System.EventHandler(this.runToEnd);
            // 
            // tabPage
            // 
            this.tabPage.Controls.Add(this.button3);
            this.tabPage.Controls.Add(this.button2);
            this.tabPage.Controls.Add(this.button1);
            this.tabPage.Controls.Add(this.richTextBox);
            this.tabPage.Location = new System.Drawing.Point(4, 25);
            this.tabPage.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage.Name = "tabPage";
            this.tabPage.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage.Size = new System.Drawing.Size(2333, 948);
            this.tabPage.TabIndex = 1;
            this.tabPage.Text = "solutions";
            this.tabPage.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(958, 330);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(216, 97);
            this.button3.TabIndex = 3;
            this.button3.Text = "neighbor Search";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.NBS);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(958, 198);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(216, 86);
            this.button2.TabIndex = 2;
            this.button2.Text = "Greedy";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.nearestInsert);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(958, 42);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(216, 101);
            this.button1.TabIndex = 1;
            this.button1.Text = "Route By Emergency Rate";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.getlowerbound);
            // 
            // richTextBox
            // 
            this.richTextBox.Location = new System.Drawing.Point(20, 42);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(879, 499);
            this.richTextBox.TabIndex = 0;
            this.richTextBox.Text = "";
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Size = new System.Drawing.Size(2333, 948);
            this.tabPage2.TabIndex = 2;
            this.tabPage2.Text = "Computed Optional Routes";
            this.tabPage2.UseVisualStyleBackColor = true;
            this.tabPage2.Paint += new System.Windows.Forms.PaintEventHandler(this.DrwaRoutes);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.Location = new System.Drawing.Point(0, 30);
            this.splitContainer1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.tabControl3);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.tabControl);
            this.splitContainer1.Size = new System.Drawing.Size(2078, 986);
            this.splitContainer1.SplitterDistance = 407;
            this.splitContainer1.SplitterWidth = 6;
            this.splitContainer1.TabIndex = 3;
            // 
            // tabControl3
            // 
            this.tabControl3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl3.Controls.Add(this.tabPage5);
            this.tabControl3.Controls.Add(this.tabPage6);
            this.tabControl3.Location = new System.Drawing.Point(4, 5);
            this.tabControl3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabControl3.Name = "tabControl3";
            this.tabControl3.SelectedIndex = 0;
            this.tabControl3.Size = new System.Drawing.Size(403, 1023);
            this.tabControl3.TabIndex = 0;
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.groupBox2);
            this.tabPage5.Controls.Add(this.groupBox1);
            this.tabPage5.Controls.Add(this.labelfilename);
            this.tabPage5.Controls.Add(this.labeltitle);
            this.tabPage5.Controls.Add(this.checkBoxGreedy);
            this.tabPage5.Location = new System.Drawing.Point(4, 34);
            this.tabPage5.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage5.Size = new System.Drawing.Size(395, 985);
            this.tabPage5.TabIndex = 0;
            this.tabPage5.Text = "Original Problem Setting";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.txbtimehorizon);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.txbcapacity);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.textBoxPnDTime);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.txbspeed);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Location = new System.Drawing.Point(0, 123);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(388, 393);
            this.groupBox2.TabIndex = 19;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Problem Setting";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(179, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "Planning  Horizon";
            // 
            // txbtimehorizon
            // 
            this.txbtimehorizon.Location = new System.Drawing.Point(193, 48);
            this.txbtimehorizon.Name = "txbtimehorizon";
            this.txbtimehorizon.Size = new System.Drawing.Size(100, 34);
            this.txbtimehorizon.TabIndex = 3;
            this.txbtimehorizon.Text = "120";
            this.txbtimehorizon.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(302, 51);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(57, 25);
            this.label7.TabIndex = 5;
            this.label7.Text = "mins";
            // 
            // txbcapacity
            // 
            this.txbcapacity.Location = new System.Drawing.Point(192, 103);
            this.txbcapacity.Name = "txbcapacity";
            this.txbcapacity.Size = new System.Drawing.Size(100, 34);
            this.txbcapacity.TabIndex = 4;
            this.txbcapacity.Text = "14";
            this.txbcapacity.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(35, 248);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(324, 25);
            this.label10.TabIndex = 13;
            this.label10.Text = "Time cost on pickup  and delivery";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(223, 322);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(96, 25);
            this.label5.TabIndex = 12;
            this.label5.Text = "per/mins";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 103);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(148, 25);
            this.label3.TabIndex = 1;
            this.label3.Text = "Truck Capacity";
            // 
            // textBoxPnDTime
            // 
            this.textBoxPnDTime.Location = new System.Drawing.Point(104, 319);
            this.textBoxPnDTime.Name = "textBoxPnDTime";
            this.textBoxPnDTime.Size = new System.Drawing.Size(100, 34);
            this.textBoxPnDTime.TabIndex = 11;
            this.textBoxPnDTime.Text = "0.25";
            this.textBoxPnDTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(302, 103);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(86, 25);
            this.label8.TabIndex = 6;
            this.label8.Text = "amount";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(99, 281);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(206, 25);
            this.label9.TabIndex = 10;
            this.label9.Text = " one common good ";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(8, 161);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(127, 25);
            this.label6.TabIndex = 2;
            this.label6.Text = "Truck Speed";
            // 
            // txbspeed
            // 
            this.txbspeed.Location = new System.Drawing.Point(193, 158);
            this.txbspeed.Name = "txbspeed";
            this.txbspeed.Size = new System.Drawing.Size(100, 34);
            this.txbspeed.TabIndex = 7;
            this.txbspeed.Text = "0.6666";
            this.txbspeed.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(293, 161);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 25);
            this.label2.TabIndex = 8;
            this.label2.Text = "km/mins";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.labelIteration);
            this.groupBox1.Controls.Add(this.checkBox_times);
            this.groupBox1.Controls.Add(this.txbCPUTime);
            this.groupBox1.Location = new System.Drawing.Point(20, 569);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(296, 152);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Stop Criterion";
            // 
            // labelIteration
            // 
            this.labelIteration.AutoSize = true;
            this.labelIteration.Location = new System.Drawing.Point(26, 109);
            this.labelIteration.Name = "labelIteration";
            this.labelIteration.Size = new System.Drawing.Size(105, 25);
            this.labelIteration.TabIndex = 20;
            this.labelIteration.Text = "Iterations:";
            // 
            // checkBox_times
            // 
            this.checkBox_times.AutoSize = true;
            this.checkBox_times.Location = new System.Drawing.Point(10, 55);
            this.checkBox_times.Name = "checkBox_times";
            this.checkBox_times.Size = new System.Drawing.Size(130, 29);
            this.checkBox_times.TabIndex = 13;
            this.checkBox_times.Text = "CPU times";
            this.checkBox_times.UseVisualStyleBackColor = true;
            // 
            // txbCPUTime
            // 
            this.txbCPUTime.Location = new System.Drawing.Point(172, 57);
            this.txbCPUTime.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txbCPUTime.Name = "txbCPUTime";
            this.txbCPUTime.Size = new System.Drawing.Size(70, 34);
            this.txbCPUTime.TabIndex = 11;
            this.txbCPUTime.Text = "2000";
            this.txbCPUTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // labelfilename
            // 
            this.labelfilename.AutoSize = true;
            this.labelfilename.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.labelfilename.Location = new System.Drawing.Point(3, 23);
            this.labelfilename.Name = "labelfilename";
            this.labelfilename.Size = new System.Drawing.Size(116, 25);
            this.labelfilename.TabIndex = 17;
            this.labelfilename.Text = "File Name :";
            // 
            // labeltitle
            // 
            this.labeltitle.AutoSize = true;
            this.labeltitle.Font = new System.Drawing.Font("微軟正黑體", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labeltitle.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.labeltitle.Location = new System.Drawing.Point(7, 72);
            this.labeltitle.Name = "labeltitle";
            this.labeltitle.Size = new System.Drawing.Size(73, 30);
            this.labeltitle.TabIndex = 14;
            this.labeltitle.Text = "Title :";
            // 
            // checkBoxGreedy
            // 
            this.checkBoxGreedy.AutoSize = true;
            this.checkBoxGreedy.Location = new System.Drawing.Point(25, 522);
            this.checkBoxGreedy.Name = "checkBoxGreedy";
            this.checkBoxGreedy.Size = new System.Drawing.Size(291, 29);
            this.checkBoxGreedy.TabIndex = 9;
            this.checkBoxGreedy.Text = "reset with greedy algorithm";
            this.checkBoxGreedy.UseVisualStyleBackColor = true;
            // 
            // tabPage6
            // 
            this.tabPage6.Location = new System.Drawing.Point(4, 25);
            this.tabPage6.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage6.Name = "tabPage6";
            this.tabPage6.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage6.Size = new System.Drawing.Size(395, 994);
            this.tabPage6.TabIndex = 1;
            this.tabPage6.Text = "tabPage6";
            this.tabPage6.UseVisualStyleBackColor = true;
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton1,
            this.toolStripButton2});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1906, 27);
            this.toolStrip1.TabIndex = 4;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(95, 24);
            this.toolStripButton1.Text = "OpenFile";
            this.toolStripButton1.Click += new System.EventHandler(this.openfile);
            // 
            // toolStripButton2
            // 
            this.toolStripButton2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton2.Image")));
            this.toolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton2.Name = "toolStripButton2";
            this.toolStripButton2.Size = new System.Drawing.Size(117, 24);
            this.toolStripButton2.Text = "SaveToExcel";
            this.toolStripButton2.Click += new System.EventHandler(this.saveToExcel);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog1";
            // 
            // checkBox
            // 
            this.checkBox.AutoSize = true;
            this.checkBox.Location = new System.Drawing.Point(291, 0);
            this.checkBox.Name = "checkBox";
            this.checkBox.Size = new System.Drawing.Size(212, 29);
            this.checkBox.TabIndex = 5;
            this.checkBox.Text = "Create Status Table";
            this.checkBox.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1906, 904);
            this.Controls.Add(this.checkBox);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.splitContainer1);
            this.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "MainForm";
            this.Text = "BikeSharingSystem";
            this.tabControl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart)).EndInit();
            this.tabPage.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.tabControl3.ResumeLayout(false);
            this.tabPage5.ResumeLayout(false);
            this.tabPage5.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart;
        private System.Windows.Forms.Button butreset;
        private System.Windows.Forms.Button butrunend;
        private System.Windows.Forms.PropertyGrid ppg;
        private System.Windows.Forms.Label labelBO;
        private System.Windows.Forms.Button butpga;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TabControl tabControl3;
        private System.Windows.Forms.TabPage tabPage5;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txbcapacity;
        private System.Windows.Forms.TextBox txbtimehorizon;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabPage tabPage6;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.ToolStripButton toolStripButton2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txbspeed;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.RichTextBox richTextBox;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label labelMax;
        private System.Windows.Forms.CheckBox checkBox;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox checkBoxGreedy;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBoxPnDTime;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label labeltitle;
        private System.Windows.Forms.Label labelfilename;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox checkBox_times;
        private System.Windows.Forms.TextBox txbCPUTime;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label labelIteration;
        private System.Windows.Forms.TabPage tabPage2;
    }
}

