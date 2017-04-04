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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.backgroundWorker2 = new System.ComponentModel.BackgroundWorker();
            this.tabPage_ServiceBull = new System.Windows.Forms.TabPage();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.button1_WrkInstuction = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.textBox_SBNum = new System.Windows.Forms.TextBox();
            this.tabControl_Main = new System.Windows.Forms.TabControl();
            this.progressBar_total = new System.Windows.Forms.ProgressBar();
            this.tabPage_ServiceBull.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tabControl_Main.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // tabPage_ServiceBull
            // 
            this.tabPage_ServiceBull.Controls.Add(this.tableLayoutPanel1);
            this.tabPage_ServiceBull.Location = new System.Drawing.Point(4, 22);
            this.tabPage_ServiceBull.Name = "tabPage_ServiceBull";
            this.tabPage_ServiceBull.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_ServiceBull.Size = new System.Drawing.Size(474, 86);
            this.tabPage_ServiceBull.TabIndex = 0;
            this.tabPage_ServiceBull.Text = "Boeing Service Bulletin";
            this.tabPage_ServiceBull.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 79F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.progressBar1, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.button1_WrkInstuction, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.panel2, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.progressBar_total, 0, 2);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 41.25F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 27.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 31.25F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(468, 80);
            this.tableLayoutPanel1.TabIndex = 3;
            // 
            // progressBar1
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.progressBar1, 2);
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progressBar1.Location = new System.Drawing.Point(3, 36);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(462, 16);
            this.progressBar1.TabIndex = 6;
            // 
            // button1_WrkInstuction
            // 
            this.button1_WrkInstuction.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button1_WrkInstuction.Location = new System.Drawing.Point(3, 3);
            this.button1_WrkInstuction.Name = "button1_WrkInstuction";
            this.button1_WrkInstuction.Size = new System.Drawing.Size(73, 27);
            this.button1_WrkInstuction.TabIndex = 3;
            this.button1_WrkInstuction.Text = "Extract";
            this.button1_WrkInstuction.UseVisualStyleBackColor = true;
            this.button1_WrkInstuction.Click += new System.EventHandler(this.button1_WrkInstuction_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.textBox_SBNum);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(82, 3);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(383, 27);
            this.panel2.TabIndex = 17;
            // 
            // textBox_SBNum
            // 
            this.textBox_SBNum.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_SBNum.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox_SBNum.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_SBNum.Location = new System.Drawing.Point(0, 0);
            this.textBox_SBNum.Name = "textBox_SBNum";
            this.textBox_SBNum.ReadOnly = true;
            this.textBox_SBNum.Size = new System.Drawing.Size(383, 23);
            this.textBox_SBNum.TabIndex = 13;
            this.textBox_SBNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // tabControl_Main
            // 
            this.tabControl_Main.Controls.Add(this.tabPage_ServiceBull);
            this.tabControl_Main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl_Main.Location = new System.Drawing.Point(0, 0);
            this.tabControl_Main.Name = "tabControl_Main";
            this.tabControl_Main.SelectedIndex = 0;
            this.tabControl_Main.Size = new System.Drawing.Size(482, 112);
            this.tabControl_Main.TabIndex = 4;
            // 
            // progressBar_total
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.progressBar_total, 2);
            this.progressBar_total.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progressBar_total.Location = new System.Drawing.Point(3, 58);
            this.progressBar_total.Name = "progressBar_total";
            this.progressBar_total.Size = new System.Drawing.Size(462, 19);
            this.progressBar_total.TabIndex = 18;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(482, 112);
            this.Controls.Add(this.tabControl_Main);
            this.MaximumSize = new System.Drawing.Size(498, 150);
            this.MinimumSize = new System.Drawing.Size(16, 132);
            this.Name = "Form1";
            this.Text = "Form1";
            this.tabPage_ServiceBull.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.tabControl_Main.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.ComponentModel.BackgroundWorker backgroundWorker2;
        private System.Windows.Forms.TabPage tabPage_ServiceBull;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button button1_WrkInstuction;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox textBox_SBNum;
        private System.Windows.Forms.TabControl tabControl_Main;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ProgressBar progressBar_total;
        
    }
}

