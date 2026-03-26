namespace KaizhongRPA
{
    partial class RpaInfoAttr
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RpaInfoAttr));
            this.lab_ClassName = new System.Windows.Forms.Label();
            this.lab_Name = new System.Windows.Forms.Label();
            this.lab_RunTime1 = new System.Windows.Forms.Label();
            this.lab_RunTime2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.txt_ClassName = new System.Windows.Forms.TextBox();
            this.txt_Name = new System.Windows.Forms.TextBox();
            this.txt_RunTime1 = new System.Windows.Forms.TextBox();
            this.lab_PathStype = new System.Windows.Forms.Label();
            this.lab_ConfigPath = new System.Windows.Forms.Label();
            this.txt_RunTime2 = new System.Windows.Forms.TextBox();
            this.txt_ConfigPath = new System.Windows.Forms.TextBox();
            this.cb_PathStype = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lab_ClassName
            // 
            this.lab_ClassName.AutoSize = true;
            this.lab_ClassName.Location = new System.Drawing.Point(26, 23);
            this.lab_ClassName.Name = "lab_ClassName";
            this.lab_ClassName.Size = new System.Drawing.Size(53, 12);
            this.lab_ClassName.TabIndex = 0;
            this.lab_ClassName.Text = "流程ID：";
            // 
            // lab_Name
            // 
            this.lab_Name.AutoSize = true;
            this.lab_Name.Location = new System.Drawing.Point(216, 23);
            this.lab_Name.Name = "lab_Name";
            this.lab_Name.Size = new System.Drawing.Size(65, 12);
            this.lab_Name.TabIndex = 1;
            this.lab_Name.Text = "流程名称：";
            // 
            // lab_RunTime1
            // 
            this.lab_RunTime1.AutoSize = true;
            this.lab_RunTime1.Location = new System.Drawing.Point(26, 105);
            this.lab_RunTime1.Name = "lab_RunTime1";
            this.lab_RunTime1.Size = new System.Drawing.Size(89, 12);
            this.lab_RunTime1.TabIndex = 2;
            this.lab_RunTime1.Text = "运行条件(从)：";
            // 
            // lab_RunTime2
            // 
            this.lab_RunTime2.AutoSize = true;
            this.lab_RunTime2.Location = new System.Drawing.Point(216, 105);
            this.lab_RunTime2.Name = "lab_RunTime2";
            this.lab_RunTime2.Size = new System.Drawing.Size(89, 12);
            this.lab_RunTime2.TabIndex = 3;
            this.lab_RunTime2.Text = "运行条件(至)：";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.Control;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(510, 378);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(79, 36);
            this.button1.TabIndex = 4;
            this.button1.Text = "修改并保存";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txt_ClassName
            // 
            this.txt_ClassName.Location = new System.Drawing.Point(28, 38);
            this.txt_ClassName.Name = "txt_ClassName";
            this.txt_ClassName.ReadOnly = true;
            this.txt_ClassName.Size = new System.Drawing.Size(160, 21);
            this.txt_ClassName.TabIndex = 5;
            // 
            // txt_Name
            // 
            this.txt_Name.Location = new System.Drawing.Point(218, 38);
            this.txt_Name.Name = "txt_Name";
            this.txt_Name.Size = new System.Drawing.Size(149, 21);
            this.txt_Name.TabIndex = 6;
            // 
            // txt_RunTime1
            // 
            this.txt_RunTime1.Location = new System.Drawing.Point(28, 120);
            this.txt_RunTime1.Name = "txt_RunTime1";
            this.txt_RunTime1.Size = new System.Drawing.Size(160, 21);
            this.txt_RunTime1.TabIndex = 7;
            // 
            // lab_PathStype
            // 
            this.lab_PathStype.AutoSize = true;
            this.lab_PathStype.Location = new System.Drawing.Point(26, 266);
            this.lab_PathStype.Name = "lab_PathStype";
            this.lab_PathStype.Size = new System.Drawing.Size(65, 12);
            this.lab_PathStype.TabIndex = 8;
            this.lab_PathStype.Text = "路径类型：";
            // 
            // lab_ConfigPath
            // 
            this.lab_ConfigPath.AutoSize = true;
            this.lab_ConfigPath.Location = new System.Drawing.Point(216, 266);
            this.lab_ConfigPath.Name = "lab_ConfigPath";
            this.lab_ConfigPath.Size = new System.Drawing.Size(65, 12);
            this.lab_ConfigPath.TabIndex = 9;
            this.lab_ConfigPath.Text = "配置文件：";
            // 
            // txt_RunTime2
            // 
            this.txt_RunTime2.Location = new System.Drawing.Point(218, 120);
            this.txt_RunTime2.Name = "txt_RunTime2";
            this.txt_RunTime2.Size = new System.Drawing.Size(149, 21);
            this.txt_RunTime2.TabIndex = 10;
            // 
            // txt_ConfigPath
            // 
            this.txt_ConfigPath.Location = new System.Drawing.Point(218, 281);
            this.txt_ConfigPath.Name = "txt_ConfigPath";
            this.txt_ConfigPath.Size = new System.Drawing.Size(332, 21);
            this.txt_ConfigPath.TabIndex = 11;
            // 
            // cb_PathStype
            // 
            this.cb_PathStype.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_PathStype.FormattingEnabled = true;
            this.cb_PathStype.Items.AddRange(new object[] {
            "相对路径",
            "绝对路径"});
            this.cb_PathStype.Location = new System.Drawing.Point(28, 281);
            this.cb_PathStype.Name = "cb_PathStype";
            this.cb_PathStype.Size = new System.Drawing.Size(160, 20);
            this.cb_PathStype.TabIndex = 12;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.OrangeRed;
            this.label1.Location = new System.Drawing.Point(26, 154);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 13;
            this.label1.Text = "label1";
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label2.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.label2.Location = new System.Drawing.Point(26, 313);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(563, 51);
            this.label2.TabIndex = 14;
            this.label2.Text = "label2";
            // 
            // RpaInfoAttr
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(631, 426);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cb_PathStype);
            this.Controls.Add(this.txt_ConfigPath);
            this.Controls.Add(this.txt_RunTime2);
            this.Controls.Add(this.lab_ConfigPath);
            this.Controls.Add(this.lab_PathStype);
            this.Controls.Add(this.txt_RunTime1);
            this.Controls.Add(this.txt_Name);
            this.Controls.Add(this.txt_ClassName);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lab_RunTime2);
            this.Controls.Add(this.lab_RunTime1);
            this.Controls.Add(this.lab_Name);
            this.Controls.Add(this.lab_ClassName);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "RpaInfoAttr";
            this.Text = "RpaInfoAttr";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lab_ClassName;
        private System.Windows.Forms.Label lab_Name;
        private System.Windows.Forms.Label lab_RunTime1;
        private System.Windows.Forms.Label lab_RunTime2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txt_ClassName;
        private System.Windows.Forms.TextBox txt_Name;
        private System.Windows.Forms.TextBox txt_RunTime1;
        private System.Windows.Forms.Label lab_PathStype;
        private System.Windows.Forms.Label lab_ConfigPath;
        private System.Windows.Forms.TextBox txt_RunTime2;
        private System.Windows.Forms.TextBox txt_ConfigPath;
        private System.Windows.Forms.ComboBox cb_PathStype;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}