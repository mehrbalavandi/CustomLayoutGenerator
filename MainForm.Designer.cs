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
            btnMergeSpeaker = new Button();
            btnApplyAndMerge = new Button();
            txtSpeakers = new TextBox();
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
            // btnMergeSpeaker
            // 
            btnMergeSpeaker.Location = new Point(12, 12);
            btnMergeSpeaker.Name = "btnMergeSpeaker";
            btnMergeSpeaker.Size = new Size(200, 23);
            btnMergeSpeaker.TabIndex = 2;
            btnMergeSpeaker.Text = "ادغام گوینده‌ها";
            btnMergeSpeaker.UseVisualStyleBackColor = true;
            btnMergeSpeaker.Click += btnMerge_Click;
            // 
            // btnApplyAndMerge
            // 
            btnApplyAndMerge.Location = new Point(12, 72);
            btnApplyAndMerge.Name = "btnApplyAndMerge";
            btnApplyAndMerge.Size = new Size(200, 23);
            btnApplyAndMerge.TabIndex = 4;
            btnApplyAndMerge.Text = "تعریف و ادغام گوینده";
            btnApplyAndMerge.UseVisualStyleBackColor = true;
            btnApplyAndMerge.Click += btnApplyAndMerge_Click;
            // 
            // txtSpeakers
            // 
            txtSpeakers.Location = new Point(12, 101);
            txtSpeakers.Name = "txtSpeakers";
            txtSpeakers.PlaceholderText = "Examiner; Candidate; Student A";
            txtSpeakers.Size = new Size(776, 23);
            txtSpeakers.TabIndex = 6;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(txtSpeakers);
            Controls.Add(btnApplyAndMerge);
            Controls.Add(btnMergeSpeaker);
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
        private Button btnMergeSpeaker;
        private Button btnApplyAndMerge;
        private TextBox txtSpeakers;
    }
}
