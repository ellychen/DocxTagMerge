namespace DocxTagFinder
{
    partial class Form1
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
            this.btnFinder = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtStartTag = new System.Windows.Forms.TextBox();
            this.txtPrefixTag = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtCloseTag = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtEndTag = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtRawFile = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtSavePath = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnFinder
            // 
            this.btnFinder.Location = new System.Drawing.Point(584, 52);
            this.btnFinder.Name = "btnFinder";
            this.btnFinder.Size = new System.Drawing.Size(117, 45);
            this.btnFinder.TabIndex = 0;
            this.btnFinder.Text = "Run";
            this.btnFinder.UseVisualStyleBackColor = true;
            this.btnFinder.Click += new System.EventHandler(this.btnFinderClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(33, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "啟始標籤";
            // 
            // txtStartTag
            // 
            this.txtStartTag.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtStartTag.Location = new System.Drawing.Point(110, 34);
            this.txtStartTag.MaxLength = 1;
            this.txtStartTag.Name = "txtStartTag";
            this.txtStartTag.Size = new System.Drawing.Size(40, 27);
            this.txtStartTag.TabIndex = 2;
            this.txtStartTag.Text = "{";
            // 
            // txtPrefixTag
            // 
            this.txtPrefixTag.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtPrefixTag.Location = new System.Drawing.Point(289, 34);
            this.txtPrefixTag.MaxLength = 10;
            this.txtPrefixTag.Name = "txtPrefixTag";
            this.txtPrefixTag.Size = new System.Drawing.Size(40, 27);
            this.txtPrefixTag.TabIndex = 4;
            this.txtPrefixTag.Text = "{{";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label2.Location = new System.Drawing.Point(180, 37);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(103, 16);
            this.label2.TabIndex = 3;
            this.label2.Text = "完整啟始標籤";
            // 
            // txtCloseTag
            // 
            this.txtCloseTag.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtCloseTag.Location = new System.Drawing.Point(289, 78);
            this.txtCloseTag.MaxLength = 10;
            this.txtCloseTag.Name = "txtCloseTag";
            this.txtCloseTag.Size = new System.Drawing.Size(40, 27);
            this.txtCloseTag.TabIndex = 8;
            this.txtCloseTag.Text = "}}";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label3.Location = new System.Drawing.Point(180, 81);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(103, 16);
            this.label3.TabIndex = 7;
            this.label3.Text = "完整結尾標籤";
            // 
            // txtEndTag
            // 
            this.txtEndTag.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtEndTag.Location = new System.Drawing.Point(110, 78);
            this.txtEndTag.MaxLength = 1;
            this.txtEndTag.Name = "txtEndTag";
            this.txtEndTag.Size = new System.Drawing.Size(40, 27);
            this.txtEndTag.TabIndex = 6;
            this.txtEndTag.Text = "}";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label4.Location = new System.Drawing.Point(33, 81);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 16);
            this.label4.TabIndex = 5;
            this.label4.Text = "最後標籤";
            // 
            // txtRawFile
            // 
            this.txtRawFile.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtRawFile.Location = new System.Drawing.Point(121, 152);
            this.txtRawFile.Name = "txtRawFile";
            this.txtRawFile.Size = new System.Drawing.Size(594, 27);
            this.txtRawFile.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label5.Location = new System.Drawing.Point(33, 154);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(71, 16);
            this.label5.TabIndex = 10;
            this.label5.Text = "原始檔案";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label6.Location = new System.Drawing.Point(33, 197);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(71, 16);
            this.label6.TabIndex = 12;
            this.label6.Text = "儲存位置";
            // 
            // txtSavePath
            // 
            this.txtSavePath.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtSavePath.Location = new System.Drawing.Point(121, 194);
            this.txtSavePath.Name = "txtSavePath";
            this.txtSavePath.Size = new System.Drawing.Size(594, 27);
            this.txtSavePath.TabIndex = 11;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 273);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtSavePath);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtRawFile);
            this.Controls.Add(this.txtCloseTag);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtEndTag);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtPrefixTag);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtStartTag);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnFinder);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Docx 標籤整理";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button btnFinder;
        private Label label1;
        private TextBox txtStartTag;
        private TextBox txtPrefixTag;
        private Label label2;
        private TextBox txtCloseTag;
        private Label label3;
        private TextBox txtEndTag;
        private Label label4;
        private TextBox txtRawFile;
        private Label label5;
        private Label label6;
        private TextBox txtSavePath;
    }
}