namespace AddInMon
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tsEditItems = new System.Windows.Forms.ToolStripMenuItem();
            this.tsMonitor = new System.Windows.Forms.ToolStripMenuItem();
            this.tsenableDebugLogToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsQuit = new System.Windows.Forms.ToolStripMenuItem();
            this.lstWord = new System.Windows.Forms.ListBox();
            this.btnClear = new System.Windows.Forms.Button();
            this.chkAutoScroll = new System.Windows.Forms.CheckBox();
            this.txtRegEx = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.lblGPO = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblVer = new System.Windows.Forms.Label();
            this.btnDebug = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnMon = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.contextMenuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.ContextMenuStrip = this.contextMenuStrip1;
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "Office Add-In Monitor";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseDoubleClick);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsEditItems,
            this.tsMonitor,
            this.tsenableDebugLogToolStripMenuItem,
            this.tsQuit});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(171, 92);
            // 
            // tsEditItems
            // 
            this.tsEditItems.Name = "tsEditItems";
            this.tsEditItems.Size = new System.Drawing.Size(170, 22);
            this.tsEditItems.Text = "&Show Window";
            this.tsEditItems.Click += new System.EventHandler(this.tsEditItems_Click);
            // 
            // tsMonitor
            // 
            this.tsMonitor.Name = "tsMonitor";
            this.tsMonitor.Size = new System.Drawing.Size(170, 22);
            this.tsMonitor.Text = "&Start Monitor";
            this.tsMonitor.Click += new System.EventHandler(this.tsMonitor_Click);
            // 
            // tsenableDebugLogToolStripMenuItem
            // 
            this.tsenableDebugLogToolStripMenuItem.Name = "tsenableDebugLogToolStripMenuItem";
            this.tsenableDebugLogToolStripMenuItem.Size = new System.Drawing.Size(170, 22);
            this.tsenableDebugLogToolStripMenuItem.Text = "Enable &Debug Log";
            this.tsenableDebugLogToolStripMenuItem.Click += new System.EventHandler(this.tsenableDebugLogToolStripMenuItem_Click);
            // 
            // tsQuit
            // 
            this.tsQuit.Name = "tsQuit";
            this.tsQuit.Size = new System.Drawing.Size(170, 22);
            this.tsQuit.Text = "E&xit Program";
            this.tsQuit.Click += new System.EventHandler(this.tsQuit_Click);
            // 
            // lstWord
            // 
            this.lstWord.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.lstWord.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstWord.FormattingEnabled = true;
            this.lstWord.Location = new System.Drawing.Point(9, 20);
            this.lstWord.Name = "lstWord";
            this.lstWord.Size = new System.Drawing.Size(661, 342);
            this.lstWord.TabIndex = 3;
            this.lstWord.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.lstWord_DrawItem);
            // 
            // btnClear
            // 
            this.btnClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClear.Location = new System.Drawing.Point(595, 368);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 23);
            this.btnClear.TabIndex = 4;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // chkAutoScroll
            // 
            this.chkAutoScroll.AutoSize = true;
            this.chkAutoScroll.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkAutoScroll.Location = new System.Drawing.Point(304, 0);
            this.chkAutoScroll.Name = "chkAutoScroll";
            this.chkAutoScroll.Size = new System.Drawing.Size(74, 17);
            this.chkAutoScroll.TabIndex = 5;
            this.chkAutoScroll.Text = "AutoScroll";
            this.chkAutoScroll.UseVisualStyleBackColor = true;
            this.chkAutoScroll.CheckedChanged += new System.EventHandler(this.chkAutoScroll_CheckedChanged);
            // 
            // txtRegEx
            // 
            this.txtRegEx.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRegEx.Location = new System.Drawing.Point(15, 19);
            this.txtRegEx.Multiline = true;
            this.txtRegEx.Name = "txtRegEx";
            this.txtRegEx.Size = new System.Drawing.Size(655, 101);
            this.txtRegEx.TabIndex = 8;
            this.txtRegEx.TextChanged += new System.EventHandler(this.txtRegEx_TextChanged);
            // 
            // btnSave
            // 
            this.btnSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSave.Location = new System.Drawing.Point(489, 126);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 9;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // lblGPO
            // 
            this.lblGPO.AutoSize = true;
            this.lblGPO.ForeColor = System.Drawing.Color.OrangeRed;
            this.lblGPO.Location = new System.Drawing.Point(471, 1);
            this.lblGPO.Name = "lblGPO";
            this.lblGPO.Size = new System.Drawing.Size(205, 13);
            this.lblGPO.TabIndex = 10;
            this.lblGPO.Text = "*Setting are being applied via GPO";
            this.lblGPO.Visible = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.lblVer);
            this.groupBox1.Controls.Add(this.btnDebug);
            this.groupBox1.Controls.Add(this.chkAutoScroll);
            this.groupBox1.Controls.Add(this.btnClear);
            this.groupBox1.Controls.Add(this.lstWord);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(4, 168);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(676, 394);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Realtime List of Active Office Add-Ins/Templates:";
            // 
            // lblVer
            // 
            this.lblVer.AutoSize = true;
            this.lblVer.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVer.Location = new System.Drawing.Point(3, 376);
            this.lblVer.Name = "lblVer";
            this.lblVer.Size = new System.Drawing.Size(35, 13);
            this.lblVer.TabIndex = 7;
            this.lblVer.Text = "label1";
            // 
            // btnDebug
            // 
            this.btnDebug.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDebug.Location = new System.Drawing.Point(514, 368);
            this.btnDebug.Name = "btnDebug";
            this.btnDebug.Size = new System.Drawing.Size(75, 23);
            this.btnDebug.TabIndex = 6;
            this.btnDebug.Text = "Open Log";
            this.btnDebug.UseVisualStyleBackColor = true;
            this.btnDebug.Click += new System.EventHandler(this.btnDebug_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnMon);
            this.groupBox2.Controls.Add(this.lblGPO);
            this.groupBox2.Controls.Add(this.txtRegEx);
            this.groupBox2.Controls.Add(this.btnSave);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(4, 7);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(676, 148);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Regular Expression used to monitor Add-Ins/Templates from getting disabled:";
            // 
            // btnMon
            // 
            this.btnMon.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMon.Location = new System.Drawing.Point(570, 126);
            this.btnMon.Name = "btnMon";
            this.btnMon.Size = new System.Drawing.Size(92, 23);
            this.btnMon.TabIndex = 11;
            this.btnMon.Text = "Start Monitor";
            this.btnMon.UseVisualStyleBackColor = true;
            this.btnMon.Click += new System.EventHandler(this.btnMon_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.panel1.Location = new System.Drawing.Point(201, 374);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(18, 16);
            this.panel1.TabIndex = 8;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(225, 376);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(173, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "= highlighted add-ins are protected.";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(686, 565);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.ShowInTaskbar = false;
            this.Text = "Office Add-In Monitor";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Resize += new System.EventHandler(this.Form1_Resize);
            this.contextMenuStrip1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tsEditItems;
        private System.Windows.Forms.ToolStripMenuItem tsMonitor;
        private System.Windows.Forms.ToolStripMenuItem tsQuit;
        private System.Windows.Forms.ListBox lstWord;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.CheckBox chkAutoScroll;
        private System.Windows.Forms.TextBox txtRegEx;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label lblGPO;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnMon;
        private System.Windows.Forms.ToolStripMenuItem tsenableDebugLogToolStripMenuItem;
        private System.Windows.Forms.Button btnDebug;
        private System.Windows.Forms.Label lblVer;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
    }
}

