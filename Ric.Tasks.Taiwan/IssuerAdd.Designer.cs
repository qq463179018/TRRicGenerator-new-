namespace Ric.Tasks
{
    partial class IssuerAdd
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
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lbRic = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tbOrgName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tbWarrantName = new System.Windows.Forms.TextBox();
            this.tbRIC = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.tbEnglishName = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.tbWarrantIssuer = new System.Windows.Forms.TextBox();
            this.tbIssuerCode = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.ErrorMsg = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.tbEnglishFullName = new System.Windows.Forms.TextBox();
            this.tbEnglishShortName = new System.Windows.Forms.TextBox();
            this.tbEnglishBriefName = new System.Windows.Forms.TextBox();
            this.tbChineseShortName = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 11F);
            this.label1.Location = new System.Drawing.Point(158, 49);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 18);
            this.label1.TabIndex = 0;
            this.label1.Text = "RIC:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 11F);
            this.label3.Location = new System.Drawing.Point(116, 77);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(76, 18);
            this.label3.TabIndex = 2;
            this.label3.Text = "權證簡稱:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Calibri", 11F);
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(46, 24);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(409, 18);
            this.label4.TabIndex = 24;
            this.label4.Text = "Can not find issuer information for below. Please input manually.";
            // 
            // lbRic
            // 
            this.lbRic.AutoSize = true;
            this.lbRic.Font = new System.Drawing.Font("Calibri", 12F);
            this.lbRic.Location = new System.Drawing.Point(50, 46);
            this.lbRic.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbRic.Name = "lbRic";
            this.lbRic.Size = new System.Drawing.Size(0, 19);
            this.lbRic.TabIndex = 4;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tbOrgName);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.tbWarrantName);
            this.groupBox1.Controls.Add(this.tbRIC);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.lbRic);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Location = new System.Drawing.Point(12, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(476, 136);
            this.groupBox1.TabIndex = 25;
            this.groupBox1.TabStop = false;
            // 
            // tbOrgName
            // 
            this.tbOrgName.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbOrgName.Font = new System.Drawing.Font("Calibri", 11F);
            this.tbOrgName.Location = new System.Drawing.Point(197, 104);
            this.tbOrgName.Name = "tbOrgName";
            this.tbOrgName.ReadOnly = true;
            this.tbOrgName.Size = new System.Drawing.Size(246, 18);
            this.tbOrgName.TabIndex = 23;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 11F);
            this.label2.Location = new System.Drawing.Point(82, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(108, 18);
            this.label2.TabIndex = 10;
            this.label2.Text = "申請機構名稱:";
            // 
            // tbWarrantName
            // 
            this.tbWarrantName.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbWarrantName.Font = new System.Drawing.Font("Calibri", 11F);
            this.tbWarrantName.Location = new System.Drawing.Point(197, 77);
            this.tbWarrantName.Name = "tbWarrantName";
            this.tbWarrantName.ReadOnly = true;
            this.tbWarrantName.Size = new System.Drawing.Size(255, 18);
            this.tbWarrantName.TabIndex = 22;
            // 
            // tbRIC
            // 
            this.tbRIC.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbRIC.Font = new System.Drawing.Font("Calibri", 11F);
            this.tbRIC.Location = new System.Drawing.Point(197, 49);
            this.tbRIC.Name = "tbRIC";
            this.tbRIC.ReadOnly = true;
            this.tbRIC.Size = new System.Drawing.Size(191, 18);
            this.tbRIC.TabIndex = 21;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.tbEnglishName);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.tbWarrantIssuer);
            this.groupBox2.Controls.Add(this.tbIssuerCode);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.ErrorMsg);
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Controls.Add(this.tbEnglishFullName);
            this.groupBox2.Controls.Add(this.tbEnglishShortName);
            this.groupBox2.Controls.Add(this.tbEnglishBriefName);
            this.groupBox2.Controls.Add(this.tbChineseShortName);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Location = new System.Drawing.Point(12, 145);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(476, 293);
            this.groupBox2.TabIndex = 26;
            this.groupBox2.TabStop = false;
            // 
            // tbEnglishName
            // 
            this.tbEnglishName.Font = new System.Drawing.Font("Calibri", 11F);
            this.tbEnglishName.Location = new System.Drawing.Point(200, 115);
            this.tbEnglishName.Name = "tbEnglishName";
            this.tbEnglishName.Size = new System.Drawing.Size(191, 25);
            this.tbEnglishName.TabIndex = 4;
            this.tbEnglishName.Tag = "EnglishName";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Calibri", 11F);
            this.label7.Location = new System.Drawing.Point(100, 118);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(93, 18);
            this.label7.TabIndex = 18;
            this.label7.Text = "EnglishName:";
            // 
            // tbWarrantIssuer
            // 
            this.tbWarrantIssuer.Font = new System.Drawing.Font("Calibri", 11F);
            this.tbWarrantIssuer.Location = new System.Drawing.Point(200, 211);
            this.tbWarrantIssuer.Name = "tbWarrantIssuer";
            this.tbWarrantIssuer.Size = new System.Drawing.Size(191, 25);
            this.tbWarrantIssuer.TabIndex = 7;
            this.tbWarrantIssuer.Tag = "WarrantIssuer";
            // 
            // tbIssuerCode
            // 
            this.tbIssuerCode.Font = new System.Drawing.Font("Calibri", 11F);
            this.tbIssuerCode.Location = new System.Drawing.Point(200, 179);
            this.tbIssuerCode.Name = "tbIssuerCode";
            this.tbIssuerCode.Size = new System.Drawing.Size(191, 25);
            this.tbIssuerCode.TabIndex = 6;
            this.tbIssuerCode.Tag = "IssuerCode";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Calibri", 11F);
            this.label5.Location = new System.Drawing.Point(95, 214);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(98, 18);
            this.label5.TabIndex = 14;
            this.label5.Text = "WarrantIssuer:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Calibri", 11F);
            this.label6.Location = new System.Drawing.Point(112, 182);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(81, 18);
            this.label6.TabIndex = 13;
            this.label6.Text = "IssuerCode:";
            // 
            // ErrorMsg
            // 
            this.ErrorMsg.AutoSize = true;
            this.ErrorMsg.Font = new System.Drawing.Font("Calibri", 11F);
            this.ErrorMsg.ForeColor = System.Drawing.Color.Red;
            this.ErrorMsg.Location = new System.Drawing.Point(19, 243);
            this.ErrorMsg.Name = "ErrorMsg";
            this.ErrorMsg.Size = new System.Drawing.Size(46, 18);
            this.ErrorMsg.TabIndex = 20;
            this.ErrorMsg.Text = "label2";
            this.ErrorMsg.Visible = false;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Calibri", 11F);
            this.button1.Location = new System.Drawing.Point(201, 265);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 25);
            this.button1.TabIndex = 19;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // tbEnglishFullName
            // 
            this.tbEnglishFullName.Font = new System.Drawing.Font("Calibri", 11F);
            this.tbEnglishFullName.Location = new System.Drawing.Point(200, 147);
            this.tbEnglishFullName.Name = "tbEnglishFullName";
            this.tbEnglishFullName.Size = new System.Drawing.Size(191, 25);
            this.tbEnglishFullName.TabIndex = 5;
            this.tbEnglishFullName.Tag = "EnglishFullName";
            // 
            // tbEnglishShortName
            // 
            this.tbEnglishShortName.Font = new System.Drawing.Font("Calibri", 11F);
            this.tbEnglishShortName.Location = new System.Drawing.Point(200, 83);
            this.tbEnglishShortName.Name = "tbEnglishShortName";
            this.tbEnglishShortName.Size = new System.Drawing.Size(191, 25);
            this.tbEnglishShortName.TabIndex = 3;
            this.tbEnglishShortName.Tag = "EnglishShortName";
            // 
            // tbEnglishBriefName
            // 
            this.tbEnglishBriefName.Font = new System.Drawing.Font("Calibri", 11F);
            this.tbEnglishBriefName.Location = new System.Drawing.Point(200, 51);
            this.tbEnglishBriefName.Name = "tbEnglishBriefName";
            this.tbEnglishBriefName.Size = new System.Drawing.Size(191, 25);
            this.tbEnglishBriefName.TabIndex = 2;
            this.tbEnglishBriefName.Tag = "EnglishBriefName";
            // 
            // tbChineseShortName
            // 
            this.tbChineseShortName.Font = new System.Drawing.Font("Calibri", 11F);
            this.tbChineseShortName.Location = new System.Drawing.Point(200, 19);
            this.tbChineseShortName.Name = "tbChineseShortName";
            this.tbChineseShortName.Size = new System.Drawing.Size(191, 25);
            this.tbChineseShortName.TabIndex = 1;
            this.tbChineseShortName.Tag = "ChineseShortName";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Calibri", 11F);
            this.label8.Location = new System.Drawing.Point(61, 22);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(132, 18);
            this.label8.TabIndex = 6;
            this.label8.Text = "ChineseShortName:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Calibri", 11F);
            this.label9.Location = new System.Drawing.Point(67, 86);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(126, 18);
            this.label9.TabIndex = 1;
            this.label9.Text = "EnglishShortName:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Calibri", 11F);
            this.label11.Location = new System.Drawing.Point(70, 54);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(123, 18);
            this.label11.TabIndex = 0;
            this.label11.Text = "EnglishBriefName:";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Calibri", 11F);
            this.label13.Location = new System.Drawing.Point(77, 150);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(116, 18);
            this.label13.TabIndex = 2;
            this.label13.Text = "EnglishFullName:";
            // 
            // IssuerAdd
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(500, 447);
            this.ControlBox = false;
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Calibri", 12F);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "IssuerAdd";
            this.Text = "New Issuer Add";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lbRic;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox tbEnglishFullName;
        private System.Windows.Forms.TextBox tbEnglishShortName;
        private System.Windows.Forms.TextBox tbEnglishBriefName;
        private System.Windows.Forms.TextBox tbChineseShortName;
        private System.Windows.Forms.TextBox tbWarrantName;
        private System.Windows.Forms.TextBox tbRIC;
        private System.Windows.Forms.Label ErrorMsg;
        private System.Windows.Forms.TextBox tbOrgName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbWarrantIssuer;
        private System.Windows.Forms.TextBox tbIssuerCode;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tbEnglishName;
        private System.Windows.Forms.Label label7;
    }
}