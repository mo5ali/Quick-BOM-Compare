namespace Quick_BOM_Compare
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.TxtAsmName = new System.Windows.Forms.TextBox();
            this.LblPathStatus = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.Labelsapstatus = new System.Windows.Forms.Label();
            this.Label3dstatus = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.LblCnsl = new System.Windows.Forms.Label();
            this.radioButton1000 = new System.Windows.Forms.RadioButton();
            this.radioButton2000 = new System.Windows.Forms.RadioButton();
            this.radioButton8000 = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(32, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(211, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Enter Assembly name without the extension";
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(827, 623);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(190, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Made by Mohamed K Ali and ChatGPT";
            // 
            // TxtAsmName
            // 
            this.TxtAsmName.Location = new System.Drawing.Point(35, 69);
            this.TxtAsmName.Name = "TxtAsmName";
            this.TxtAsmName.Size = new System.Drawing.Size(198, 20);
            this.TxtAsmName.TabIndex = 2;
            // 
            // LblPathStatus
            // 
            this.LblPathStatus.AutoSize = true;
            this.LblPathStatus.Location = new System.Drawing.Point(35, 168);
            this.LblPathStatus.Name = "LblPathStatus";
            this.LblPathStatus.Size = new System.Drawing.Size(10, 13);
            this.LblPathStatus.TabIndex = 3;
            this.LblPathStatus.Text = "__";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(35, 238);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(10, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "__";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(35, 402);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 5;
            this.button1.Text = "Check";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(140, 402);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 6;
            this.button2.Text = "Extract";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(348, 402);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 7;
            this.button3.Text = "Export";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // Labelsapstatus
            // 
            this.Labelsapstatus.AutoSize = true;
            this.Labelsapstatus.Location = new System.Drawing.Point(35, 449);
            this.Labelsapstatus.Name = "Labelsapstatus";
            this.Labelsapstatus.Size = new System.Drawing.Size(10, 13);
            this.Labelsapstatus.TabIndex = 8;
            this.Labelsapstatus.Text = " ";
            // 
            // Label3dstatus
            // 
            this.Label3dstatus.AutoSize = true;
            this.Label3dstatus.Location = new System.Drawing.Point(32, 507);
            this.Label3dstatus.Name = "Label3dstatus";
            this.Label3dstatus.Size = new System.Drawing.Size(10, 13);
            this.Label3dstatus.TabIndex = 9;
            this.Label3dstatus.Text = " ";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(246, 402);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 23);
            this.button4.TabIndex = 10;
            this.button4.Text = "View";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(453, 22);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(539, 589);
            this.dataGridView1.TabIndex = 11;
            // 
            // LblCnsl
            // 
            this.LblCnsl.AutoSize = true;
            this.LblCnsl.Location = new System.Drawing.Point(12, 555);
            this.LblCnsl.Name = "LblCnsl";
            this.LblCnsl.Size = new System.Drawing.Size(50, 13);
            this.LblCnsl.TabIndex = 12;
            this.LblCnsl.Text = "console:-";
            // 
            // radioButton1000
            // 
            this.radioButton1000.AutoSize = true;
            this.radioButton1000.Location = new System.Drawing.Point(35, 367);
            this.radioButton1000.Name = "radioButton1000";
            this.radioButton1000.Size = new System.Drawing.Size(49, 17);
            this.radioButton1000.TabIndex = 13;
            this.radioButton1000.TabStop = true;
            this.radioButton1000.Text = "1000";
            this.radioButton1000.UseVisualStyleBackColor = true;
            // 
            // radioButton2000
            // 
            this.radioButton2000.AutoSize = true;
            this.radioButton2000.Location = new System.Drawing.Point(194, 367);
            this.radioButton2000.Name = "radioButton2000";
            this.radioButton2000.Size = new System.Drawing.Size(49, 17);
            this.radioButton2000.TabIndex = 16;
            this.radioButton2000.TabStop = true;
            this.radioButton2000.Text = "2000";
            this.radioButton2000.UseVisualStyleBackColor = true;
            // 
            // radioButton8000
            // 
            this.radioButton8000.AutoSize = true;
            this.radioButton8000.Location = new System.Drawing.Point(374, 367);
            this.radioButton8000.Name = "radioButton8000";
            this.radioButton8000.Size = new System.Drawing.Size(49, 17);
            this.radioButton8000.TabIndex = 17;
            this.radioButton8000.TabStop = true;
            this.radioButton8000.Text = "8000";
            this.radioButton8000.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1017, 634);
            this.Controls.Add(this.radioButton8000);
            this.Controls.Add(this.radioButton2000);
            this.Controls.Add(this.radioButton1000);
            this.Controls.Add(this.LblCnsl);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.Label3dstatus);
            this.Controls.Add(this.Labelsapstatus);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.LblPathStatus);
            this.Controls.Add(this.TxtAsmName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TxtAsmName;
        private System.Windows.Forms.Label LblPathStatus;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label Labelsapstatus;
        private System.Windows.Forms.Label Label3dstatus;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label LblCnsl;
        private System.Windows.Forms.RadioButton radioButton1000;
        private System.Windows.Forms.RadioButton radioButton2000;
        private System.Windows.Forms.RadioButton radioButton8000;
    }
}

