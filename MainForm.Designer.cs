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
            button1 = new Button();
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
            // button1
            // 
            button1.Location = new Point(311, 417);
            button1.Name = "button1";
            button1.Size = new Size(200, 23);
            button1.TabIndex = 2;
            button1.Text = "Indent Speaker Labels 2";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(button1);
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
        private Button button1;
    }
}
