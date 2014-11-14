namespace LogToChart
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.buttonConvert = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.radioButtonXlsx = new System.Windows.Forms.RadioButton();
            this.radioButtonXls = new System.Windows.Forms.RadioButton();
            this.radioButtonTxt = new System.Windows.Forms.RadioButton();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBoxInterval = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(135, 32);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(180, 21);
            this.textBox1.TabIndex = 0;
            this.textBox1.Text = "C:\\g5w2562g.txt";
            // 
            // buttonConvert
            // 
            this.buttonConvert.Location = new System.Drawing.Point(400, 68);
            this.buttonConvert.Name = "buttonConvert";
            this.buttonConvert.Size = new System.Drawing.Size(111, 23);
            this.buttonConvert.TabIndex = 1;
            this.buttonConvert.Text = "Convert";
            this.buttonConvert.UseVisualStyleBackColor = true;
            this.buttonConvert.Click += new System.EventHandler(this.buttonConvert_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(54, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "Source";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(54, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "Destination";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(135, 71);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(180, 21);
            this.textBox2.TabIndex = 0;
            this.textBox2.Text = "C:\\";
            // 
            // radioButtonXlsx
            // 
            this.radioButtonXlsx.AutoSize = true;
            this.radioButtonXlsx.Checked = true;
            this.radioButtonXlsx.Location = new System.Drawing.Point(331, 31);
            this.radioButtonXlsx.Name = "radioButtonXlsx";
            this.radioButtonXlsx.Size = new System.Drawing.Size(47, 16);
            this.radioButtonXlsx.TabIndex = 4;
            this.radioButtonXlsx.TabStop = true;
            this.radioButtonXlsx.Text = "xlsx";
            this.radioButtonXlsx.UseVisualStyleBackColor = true;
            // 
            // radioButtonXls
            // 
            this.radioButtonXls.AutoSize = true;
            this.radioButtonXls.Location = new System.Drawing.Point(331, 51);
            this.radioButtonXls.Name = "radioButtonXls";
            this.radioButtonXls.Size = new System.Drawing.Size(41, 16);
            this.radioButtonXls.TabIndex = 4;
            this.radioButtonXls.Text = "xls";
            this.radioButtonXls.UseVisualStyleBackColor = true;
            // 
            // radioButtonTxt
            // 
            this.radioButtonTxt.AutoSize = true;
            this.radioButtonTxt.Location = new System.Drawing.Point(331, 71);
            this.radioButtonTxt.Name = "radioButtonTxt";
            this.radioButtonTxt.Size = new System.Drawing.Size(41, 16);
            this.radioButtonTxt.TabIndex = 4;
            this.radioButtonTxt.Text = "txt";
            this.radioButtonTxt.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(398, 28);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(113, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "Time Interval(min)";
            // 
            // comboBoxInterval
            // 
            this.comboBoxInterval.FormattingEnabled = true;
            this.comboBoxInterval.Location = new System.Drawing.Point(400, 44);
            this.comboBoxInterval.Name = "comboBoxInterval";
            this.comboBoxInterval.Size = new System.Drawing.Size(111, 20);
            this.comboBoxInterval.TabIndex = 8;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(556, 131);
            this.Controls.Add(this.comboBoxInterval);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.radioButtonTxt);
            this.Controls.Add(this.radioButtonXls);
            this.Controls.Add(this.radioButtonXlsx);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonConvert);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Name = "Form1";
            this.Text = "LogToData   CopyRight@TeamNeptune.AllRightsReserved";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button buttonConvert;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.RadioButton radioButtonXlsx;
        private System.Windows.Forms.RadioButton radioButtonXls;
        private System.Windows.Forms.RadioButton radioButtonTxt;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBoxInterval;
    }
}

