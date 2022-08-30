namespace WindowsFormsApp1
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.uiButton6 = new Sunny.UI.UIButton();
            this.uiTextBox5 = new Sunny.UI.UITextBox();
            this.uiTextBox4 = new Sunny.UI.UITextBox();
            this.uiTextBox3 = new Sunny.UI.UITextBox();
            this.uiTextBox2 = new Sunny.UI.UITextBox();
            this.uiTextBox1 = new Sunny.UI.UITextBox();
            this.uiButton5 = new Sunny.UI.UIButton();
            this.uiButton4 = new Sunny.UI.UIButton();
            this.uiButton3 = new Sunny.UI.UIButton();
            this.uiButton2 = new Sunny.UI.UIButton();
            this.uiButton1 = new Sunny.UI.UIButton();
            this.label3 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
            this.保存记录ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.退出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem6 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem7 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem8 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem9 = new System.Windows.Forms.ToolStripMenuItem();
            this.合并为单个文件选择ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.标准化ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.保留首行表头ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.批量替换ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.uiStyleManager1 = new Sunny.UI.UIStyleManager(this.components);
            this.uiRichTextBox1 = new Sunny.UI.UIRichTextBox();
            this.panel1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 12);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 27);
            this.label1.TabIndex = 5;
            this.label1.Text = "表头行数";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 50);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 27);
            this.label2.TabIndex = 6;
            this.label2.Text = "拆分字段";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.uiButton6);
            this.panel1.Controls.Add(this.uiTextBox5);
            this.panel1.Controls.Add(this.uiTextBox4);
            this.panel1.Controls.Add(this.uiTextBox3);
            this.panel1.Controls.Add(this.uiTextBox2);
            this.panel1.Controls.Add(this.uiTextBox1);
            this.panel1.Controls.Add(this.uiButton5);
            this.panel1.Controls.Add(this.uiButton4);
            this.panel1.Controls.Add(this.uiButton3);
            this.panel1.Controls.Add(this.uiButton2);
            this.panel1.Controls.Add(this.uiButton1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Location = new System.Drawing.Point(4, 70);
            this.panel1.Margin = new System.Windows.Forms.Padding(4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(805, 127);
            this.panel1.TabIndex = 15;
            // 
            // uiButton6
            // 
            this.uiButton6.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uiButton6.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiButton6.Location = new System.Drawing.Point(702, 85);
            this.uiButton6.MinimumSize = new System.Drawing.Size(1, 1);
            this.uiButton6.Name = "uiButton6";
            this.uiButton6.Size = new System.Drawing.Size(100, 35);
            this.uiButton6.TabIndex = 22;
            this.uiButton6.Text = "取消";
            this.uiButton6.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiButton6.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uiButton6.Click += new System.EventHandler(this.button6_Click);
            // 
            // uiTextBox5
            // 
            this.uiTextBox5.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.uiTextBox5.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiTextBox5.Location = new System.Drawing.Point(109, 85);
            this.uiTextBox5.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.uiTextBox5.MinimumSize = new System.Drawing.Size(1, 16);
            this.uiTextBox5.Name = "uiTextBox5";
            this.uiTextBox5.ShowText = false;
            this.uiTextBox5.Size = new System.Drawing.Size(89, 35);
            this.uiTextBox5.TabIndex = 27;
            this.uiTextBox5.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiTextBox5.Watermark = "";
            this.uiTextBox5.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uiTextBox5.TextChanged += new System.EventHandler(this.textBox5_TextChanged);
            // 
            // uiTextBox4
            // 
            this.uiTextBox4.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.uiTextBox4.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiTextBox4.Location = new System.Drawing.Point(205, 45);
            this.uiTextBox4.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.uiTextBox4.MinimumSize = new System.Drawing.Size(1, 16);
            this.uiTextBox4.Name = "uiTextBox4";
            this.uiTextBox4.ShowText = false;
            this.uiTextBox4.Size = new System.Drawing.Size(490, 35);
            this.uiTextBox4.TabIndex = 26;
            this.uiTextBox4.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiTextBox4.Watermark = "";
            this.uiTextBox4.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uiTextBox4.TextChanged += new System.EventHandler(this.textBox4_TextChanged);
            // 
            // uiTextBox3
            // 
            this.uiTextBox3.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.uiTextBox3.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiTextBox3.Location = new System.Drawing.Point(205, 4);
            this.uiTextBox3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.uiTextBox3.MinimumSize = new System.Drawing.Size(1, 16);
            this.uiTextBox3.Name = "uiTextBox3";
            this.uiTextBox3.ShowText = false;
            this.uiTextBox3.Size = new System.Drawing.Size(490, 35);
            this.uiTextBox3.TabIndex = 25;
            this.uiTextBox3.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiTextBox3.Watermark = "";
            this.uiTextBox3.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uiTextBox3.TextChanged += new System.EventHandler(this.uiTextBox3_TextChanged);
            // 
            // uiTextBox2
            // 
            this.uiTextBox2.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.uiTextBox2.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiTextBox2.Location = new System.Drawing.Point(109, 45);
            this.uiTextBox2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.uiTextBox2.MinimumSize = new System.Drawing.Size(1, 16);
            this.uiTextBox2.Name = "uiTextBox2";
            this.uiTextBox2.ShowText = false;
            this.uiTextBox2.Size = new System.Drawing.Size(89, 35);
            this.uiTextBox2.TabIndex = 24;
            this.uiTextBox2.Text = "管道名称";
            this.uiTextBox2.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiTextBox2.Watermark = "";
            this.uiTextBox2.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // uiTextBox1
            // 
            this.uiTextBox1.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.uiTextBox1.DoubleValue = 4D;
            this.uiTextBox1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiTextBox1.IntValue = 4;
            this.uiTextBox1.Location = new System.Drawing.Point(109, 5);
            this.uiTextBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.uiTextBox1.MinimumSize = new System.Drawing.Size(1, 16);
            this.uiTextBox1.Name = "uiTextBox1";
            this.uiTextBox1.ShowText = false;
            this.uiTextBox1.Size = new System.Drawing.Size(89, 35);
            this.uiTextBox1.TabIndex = 23;
            this.uiTextBox1.Text = "4";
            this.uiTextBox1.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiTextBox1.Watermark = "";
            this.uiTextBox1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uiTextBox1.TextChanged += new System.EventHandler(this.uiTextBox1_TextChanged);
            // 
            // uiButton5
            // 
            this.uiButton5.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uiButton5.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiButton5.Location = new System.Drawing.Point(328, 85);
            this.uiButton5.MinimumSize = new System.Drawing.Size(1, 1);
            this.uiButton5.Name = "uiButton5";
            this.uiButton5.Size = new System.Drawing.Size(100, 35);
            this.uiButton5.TabIndex = 21;
            this.uiButton5.Text = "批量拆分";
            this.uiButton5.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiButton5.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uiButton5.Click += new System.EventHandler(this.button5_Click);
            // 
            // uiButton4
            // 
            this.uiButton4.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uiButton4.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiButton4.Location = new System.Drawing.Point(434, 85);
            this.uiButton4.MinimumSize = new System.Drawing.Size(1, 1);
            this.uiButton4.Name = "uiButton4";
            this.uiButton4.Size = new System.Drawing.Size(100, 35);
            this.uiButton4.TabIndex = 20;
            this.uiButton4.Text = "批量合并";
            this.uiButton4.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiButton4.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uiButton4.Click += new System.EventHandler(this.button4_Click);
            // 
            // uiButton3
            // 
            this.uiButton3.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uiButton3.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiButton3.Location = new System.Drawing.Point(702, 45);
            this.uiButton3.MinimumSize = new System.Drawing.Size(1, 1);
            this.uiButton3.Name = "uiButton3";
            this.uiButton3.Size = new System.Drawing.Size(100, 35);
            this.uiButton3.TabIndex = 19;
            this.uiButton3.Text = "保存";
            this.uiButton3.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiButton3.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uiButton3.Click += new System.EventHandler(this.button3_Click);
            // 
            // uiButton2
            // 
            this.uiButton2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uiButton2.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiButton2.Location = new System.Drawing.Point(205, 85);
            this.uiButton2.MinimumSize = new System.Drawing.Size(1, 1);
            this.uiButton2.Name = "uiButton2";
            this.uiButton2.Size = new System.Drawing.Size(117, 35);
            this.uiButton2.TabIndex = 18;
            this.uiButton2.Text = "单文件拆分";
            this.uiButton2.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiButton2.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uiButton2.Click += new System.EventHandler(this.button2_Click);
            // 
            // uiButton1
            // 
            this.uiButton1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uiButton1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiButton1.Location = new System.Drawing.Point(702, 4);
            this.uiButton1.MinimumSize = new System.Drawing.Size(1, 1);
            this.uiButton1.Name = "uiButton1";
            this.uiButton1.Size = new System.Drawing.Size(100, 35);
            this.uiButton1.TabIndex = 17;
            this.uiButton1.Text = "打开";
            this.uiButton1.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiButton1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uiButton1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 89);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 27);
            this.label3.TabIndex = 16;
            this.label3.Text = "文件过滤";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.menuStrip1.AutoSize = false;
            this.menuStrip1.BackColor = System.Drawing.Color.White;
            this.menuStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1,
            this.toolStripMenuItem2,
            this.toolStripMenuItem3,
            this.标准化ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(5, 40);
            this.menuStrip1.Margin = new System.Windows.Forms.Padding(1);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.menuStrip1.Size = new System.Drawing.Size(804, 31);
            this.menuStrip1.TabIndex = 16;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.BackColor = System.Drawing.Color.White;
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem4,
            this.toolStripMenuItem5,
            this.保存记录ToolStripMenuItem,
            this.退出ToolStripMenuItem});
            this.toolStripMenuItem1.Font = new System.Drawing.Font("Microsoft YaHei UI", 10F);
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(58, 27);
            this.toolStripMenuItem1.Text = "开始";
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.BackColor = System.Drawing.SystemColors.Control;
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(162, 28);
            this.toolStripMenuItem4.Text = "打开";
            this.toolStripMenuItem4.Click += new System.EventHandler(this.button1_Click);
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            this.toolStripMenuItem5.Size = new System.Drawing.Size(162, 28);
            this.toolStripMenuItem5.Text = "保存";
            this.toolStripMenuItem5.Click += new System.EventHandler(this.button3_Click);
            // 
            // 保存记录ToolStripMenuItem
            // 
            this.保存记录ToolStripMenuItem.Name = "保存记录ToolStripMenuItem";
            this.保存记录ToolStripMenuItem.Size = new System.Drawing.Size(162, 28);
            this.保存记录ToolStripMenuItem.Text = "保存记录";
            this.保存记录ToolStripMenuItem.Click += new System.EventHandler(this.保存记录ToolStripMenuItem_Click);
            // 
            // 退出ToolStripMenuItem
            // 
            this.退出ToolStripMenuItem.Name = "退出ToolStripMenuItem";
            this.退出ToolStripMenuItem.Size = new System.Drawing.Size(162, 28);
            this.退出ToolStripMenuItem.Text = "退出";
            this.退出ToolStripMenuItem.Click += new System.EventHandler(this.退出ToolStripMenuItem_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem6,
            this.toolStripMenuItem7});
            this.toolStripMenuItem2.Font = new System.Drawing.Font("Microsoft YaHei UI", 10F);
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(58, 27);
            this.toolStripMenuItem2.Text = "拆分";
            // 
            // toolStripMenuItem6
            // 
            this.toolStripMenuItem6.Name = "toolStripMenuItem6";
            this.toolStripMenuItem6.Size = new System.Drawing.Size(179, 28);
            this.toolStripMenuItem6.Text = "单文件拆分";
            this.toolStripMenuItem6.Click += new System.EventHandler(this.button2_Click);
            // 
            // toolStripMenuItem7
            // 
            this.toolStripMenuItem7.Name = "toolStripMenuItem7";
            this.toolStripMenuItem7.Size = new System.Drawing.Size(179, 28);
            this.toolStripMenuItem7.Text = "批量拆分";
            this.toolStripMenuItem7.Click += new System.EventHandler(this.button5_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem8,
            this.toolStripMenuItem9,
            this.合并为单个文件选择ToolStripMenuItem});
            this.toolStripMenuItem3.Font = new System.Drawing.Font("Microsoft YaHei UI", 10F);
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(58, 27);
            this.toolStripMenuItem3.Text = "合并";
            // 
            // toolStripMenuItem8
            // 
            this.toolStripMenuItem8.Name = "toolStripMenuItem8";
            this.toolStripMenuItem8.Size = new System.Drawing.Size(298, 28);
            this.toolStripMenuItem8.Text = "批量合并";
            this.toolStripMenuItem8.Click += new System.EventHandler(this.button4_Click);
            // 
            // toolStripMenuItem9
            // 
            this.toolStripMenuItem9.Name = "toolStripMenuItem9";
            this.toolStripMenuItem9.Size = new System.Drawing.Size(298, 28);
            this.toolStripMenuItem9.Text = "合并为单个文件（文件夹）";
            this.toolStripMenuItem9.Click += new System.EventHandler(this.toolStripMenuItem9_Click);
            // 
            // 合并为单个文件选择ToolStripMenuItem
            // 
            this.合并为单个文件选择ToolStripMenuItem.Name = "合并为单个文件选择ToolStripMenuItem";
            this.合并为单个文件选择ToolStripMenuItem.Size = new System.Drawing.Size(298, 28);
            this.合并为单个文件选择ToolStripMenuItem.Text = "合并为单个文件（选择）";
            this.合并为单个文件选择ToolStripMenuItem.Click += new System.EventHandler(this.合并为单个文件选择ToolStripMenuItem_Click);
            // 
            // 标准化ToolStripMenuItem
            // 
            this.标准化ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.保留首行表头ToolStripMenuItem,
            this.批量替换ToolStripMenuItem,
            this.删除ToolStripMenuItem});
            this.标准化ToolStripMenuItem.Font = new System.Drawing.Font("Microsoft YaHei UI", 10F);
            this.标准化ToolStripMenuItem.Name = "标准化ToolStripMenuItem";
            this.标准化ToolStripMenuItem.Size = new System.Drawing.Size(75, 27);
            this.标准化ToolStripMenuItem.Text = "标准化";
            // 
            // 保留首行表头ToolStripMenuItem
            // 
            this.保留首行表头ToolStripMenuItem.Name = "保留首行表头ToolStripMenuItem";
            this.保留首行表头ToolStripMenuItem.Size = new System.Drawing.Size(246, 28);
            this.保留首行表头ToolStripMenuItem.Text = "保留首行表头转CSV";
            this.保留首行表头ToolStripMenuItem.Click += new System.EventHandler(this.保留首行表头ToolStripMenuItem_Click);
            // 
            // 批量替换ToolStripMenuItem
            // 
            this.批量替换ToolStripMenuItem.Name = "批量替换ToolStripMenuItem";
            this.批量替换ToolStripMenuItem.Size = new System.Drawing.Size(246, 28);
            this.批量替换ToolStripMenuItem.Text = "批量替换";
            this.批量替换ToolStripMenuItem.Click += new System.EventHandler(this.批量替换ToolStripMenuItem_Click);
            // 
            // 删除ToolStripMenuItem
            // 
            this.删除ToolStripMenuItem.Name = "删除ToolStripMenuItem";
            this.删除ToolStripMenuItem.Size = new System.Drawing.Size(246, 28);
            this.删除ToolStripMenuItem.Text = "删除";
            this.删除ToolStripMenuItem.Click += new System.EventHandler(this.button7_Click);
            // 
            // uiStyleManager1
            // 
            this.uiStyleManager1.DPIScale = true;
            // 
            // uiRichTextBox1
            // 
            this.uiRichTextBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.uiRichTextBox1.FillColor = System.Drawing.Color.White;
            this.uiRichTextBox1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiRichTextBox1.Location = new System.Drawing.Point(4, 198);
            this.uiRichTextBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.uiRichTextBox1.MinimumSize = new System.Drawing.Size(1, 1);
            this.uiRichTextBox1.Name = "uiRichTextBox1";
            this.uiRichTextBox1.Padding = new System.Windows.Forms.Padding(2);
            this.uiRichTextBox1.ShowText = false;
            this.uiRichTextBox1.Size = new System.Drawing.Size(805, 494);
            this.uiRichTextBox1.TabIndex = 17;
            this.uiRichTextBox1.Text = resources.GetString("uiRichTextBox1.Text");
            this.uiRichTextBox1.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            this.uiRichTextBox1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uiRichTextBox1.TextChanged += new System.EventHandler(this.richTextBox1_TextChanged);
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(815, 698);
            this.Controls.Add(this.uiRichTextBox1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimumSize = new System.Drawing.Size(746, 488);
            this.Name = "Form1";
            this.Text = "表格处理工具 v2.0";
            this.ZoomScaleRect = new System.Drawing.Rectangle(19, 19, 746, 488);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem4;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem5;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem6;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem7;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem8;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem9;
        private System.Windows.Forms.ToolStripMenuItem 退出ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 合并为单个文件选择ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 保存记录ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 标准化ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 保留首行表头ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 批量替换ToolStripMenuItem;
        private Sunny.UI.UIStyleManager uiStyleManager1;
        private Sunny.UI.UIButton uiButton1;
        private Sunny.UI.UIRichTextBox uiRichTextBox1;
        private Sunny.UI.UIButton uiButton2;
        private Sunny.UI.UIButton uiButton6;
        private Sunny.UI.UIButton uiButton5;
        private Sunny.UI.UIButton uiButton4;
        private Sunny.UI.UIButton uiButton3;
        private System.Windows.Forms.ToolStripMenuItem 删除ToolStripMenuItem;
        private Sunny.UI.UITextBox uiTextBox2;
        private Sunny.UI.UITextBox uiTextBox1;
        private Sunny.UI.UITextBox uiTextBox5;
        private Sunny.UI.UITextBox uiTextBox4;
        private Sunny.UI.UITextBox uiTextBox3;
        public System.Windows.Forms.Panel panel1;
    }
}

