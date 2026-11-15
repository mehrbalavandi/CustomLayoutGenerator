namespace CustomLayoutGenerator
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnSelectFile = new Button();
            btnIndentLabels = new Button();
            btnMerge = new Button();
            btnApplyAndMerge = new Button();
            btnJoin = new Button();
            chkJoinRelaxed = new CheckBox();
            btnFixCyrillic = new Button();
            txtSpeakers = new RichTextBox();
            label1 = new Label();
            SuspendLayout();
            // 
            // btnSelectFile
            // 
            btnSelectFile.Location = new Point(12, 415);
            btnSelectFile.Name = "btnSelectFile";
            btnSelectFile.Size = new Size(75, 23);
            btnSelectFile.TabIndex = 0;
            btnSelectFile.Text = "Select Word File";
            btnSelectFile.UseVisualStyleBackColor = true;
            btnSelectFile.Click += btnSelectFile_Click;
            // 
            // btnIndentLabels
            // 
            btnIndentLabels.AutoSize = true;
            btnIndentLabels.Location = new Point(93, 415);
            btnIndentLabels.Name = "btnIndentLabels";
            btnIndentLabels.Size = new Size(212, 25);
            btnIndentLabels.TabIndex = 1;
            btnIndentLabels.Text = "Indent Speaker Labels (Audio Scripts)";
            btnIndentLabels.UseVisualStyleBackColor = true;
            btnIndentLabels.Click += btnIndentLabels_Click;
            // 
            // btnMerge
            // 
            btnMerge.Location = new Point(12, 273);
            btnMerge.Name = "btnMerge";
            btnMerge.Size = new Size(200, 23);
            btnMerge.TabIndex = 2;
            btnMerge.Text = "ادغام گوینده‌های دارای استایل";
            btnMerge.UseVisualStyleBackColor = true;
            btnMerge.Click += btnMerge_Click;
            // 
            // btnApplyAndMerge
            // 
            btnApplyAndMerge.Location = new Point(12, 200);
            btnApplyAndMerge.Name = "btnApplyAndMerge";
            btnApplyAndMerge.Size = new Size(200, 23);
            btnApplyAndMerge.TabIndex = 4;
            btnApplyAndMerge.Text = "اعمال استایل و ادغام گوینده‌ها";
            btnApplyAndMerge.UseVisualStyleBackColor = true;
            btnApplyAndMerge.Click += btnApplyAndMerge_Click;
            // 
            // btnJoin
            // 
            btnJoin.AutoSize = true;
            btnJoin.Location = new Point(12, 169);
            btnJoin.Name = "btnJoin";
            btnJoin.Size = new Size(223, 25);
            btnJoin.TabIndex = 7;
            btnJoin.Text = "ادغام پاراگراف‌های شکسته شده";
            btnJoin.UseVisualStyleBackColor = true;
            btnJoin.Click += btnJoin_Click;
            // 
            // chkJoinRelaxed
            // 
            chkJoinRelaxed.AutoSize = true;
            chkJoinRelaxed.Location = new Point(241, 173);
            chkJoinRelaxed.Name = "chkJoinRelaxed";
            chkJoinRelaxed.Size = new Size(156, 19);
            chkJoinRelaxed.TabIndex = 8;
            chkJoinRelaxed.Text = "حالت آزاد (برای متن فارسی)";
            chkJoinRelaxed.UseVisualStyleBackColor = true;
            // 
            // btnFixCyrillic
            // 
            btnFixCyrillic.Location = new Point(12, 12);
            btnFixCyrillic.Name = "btnFixCyrillic";
            btnFixCyrillic.Size = new Size(200, 23);
            btnFixCyrillic.TabIndex = 9;
            btnFixCyrillic.Text = "اصلاح حروف سیریلیک";
            btnFixCyrillic.UseVisualStyleBackColor = true;
            btnFixCyrillic.Click += btnFixCyrillic_Click;
            // 
            // txtSpeakers
            // 
            txtSpeakers.Location = new Point(12, 67);
            txtSpeakers.Name = "txtSpeakers";
            txtSpeakers.Size = new Size(776, 96);
            txtSpeakers.TabIndex = 10;
            txtSpeakers.Text = "";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 49);
            label1.Name = "label1";
            label1.Size = new Size(188, 15);
            label1.TabIndex = 11;
            label1.Text = "Examiner; Candidate; Student A, ...";
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(label1);
            Controls.Add(txtSpeakers);
            Controls.Add(btnFixCyrillic);
            Controls.Add(chkJoinRelaxed);
            Controls.Add(btnJoin);
            Controls.Add(btnApplyAndMerge);
            Controls.Add(btnMerge);
            Controls.Add(btnIndentLabels);
            Controls.Add(btnSelectFile);
            Name = "MainForm";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnSelectFile;
        private Button btnIndentLabels;
        private Button btnMerge;
        private Button btnApplyAndMerge;
        private Button btnJoin;
        private CheckBox chkJoinRelaxed;
        private Button btnFixCyrillic;
        private RichTextBox txtSpeakers;
        private Label label1;
    }
}
