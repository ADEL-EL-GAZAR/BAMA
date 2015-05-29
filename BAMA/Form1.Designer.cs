namespace BAMA
{
    partial class FormBAMA
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
            this.btnAnalysis = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtWord = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnAnalysis
            // 
            this.btnAnalysis.Location = new System.Drawing.Point(94, 120);
            this.btnAnalysis.Name = "btnAnalysis";
            this.btnAnalysis.Size = new System.Drawing.Size(75, 23);
            this.btnAnalysis.TabIndex = 0;
            this.btnAnalysis.Text = "Analysis";
            this.btnAnalysis.UseVisualStyleBackColor = true;
            this.btnAnalysis.Click += new System.EventHandler(this.btnAnalysis_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(202, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Enter Word To Analyse : ex (wbAlErbyp)";
            // 
            // txtWord
            // 
            this.txtWord.Location = new System.Drawing.Point(16, 40);
            this.txtWord.Name = "txtWord";
            this.txtWord.Size = new System.Drawing.Size(256, 20);
            this.txtWord.TabIndex = 2;
            // 
            // FormBAMA
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.txtWord);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAnalysis);
            this.Name = "FormBAMA";
            this.Text = "BAMA";
            this.Load += new System.EventHandler(this.FormBAMA_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAnalysis;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtWord;
    }
}

