namespace ExcelVTEntegrasyonProjesi
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            btnVTdenOku = new Button();
            richTextBox1 = new RichTextBox();
            btnExceldenOku = new Button();
            richTextBox2 = new RichTextBox();
            SuspendLayout();
            // 
            // btnVTdenOku
            // 
            btnVTdenOku.Location = new Point(556, 119);
            btnVTdenOku.Name = "btnVTdenOku";
            btnVTdenOku.Size = new Size(216, 92);
            btnVTdenOku.TabIndex = 0;
            btnVTdenOku.Text = "Veri Tabanından Oku ve Excel'e Yaz";
            btnVTdenOku.UseVisualStyleBackColor = true;
            btnVTdenOku.Click += btnVTdenOku_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(12, 119);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(521, 136);
            richTextBox1.TabIndex = 1;
            richTextBox1.Text = "";
            // 
            // btnExceldenOku
            // 
            btnExceldenOku.Location = new Point(556, 302);
            btnExceldenOku.Name = "btnExceldenOku";
            btnExceldenOku.Size = new Size(216, 92);
            btnExceldenOku.TabIndex = 3;
            btnExceldenOku.Text = "Excel'den Oku Veri Tabanına Yaz";
            btnExceldenOku.UseVisualStyleBackColor = true;
            btnExceldenOku.Click += btnExceldenOku_Click;
            // 
            // richTextBox2
            // 
            richTextBox2.Location = new Point(12, 319);
            richTextBox2.Name = "richTextBox2";
            richTextBox2.Size = new Size(521, 136);
            richTextBox2.TabIndex = 4;
            richTextBox2.Text = "";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(64, 64, 64);
            ClientSize = new Size(886, 565);
            Controls.Add(richTextBox2);
            Controls.Add(btnExceldenOku);
            Controls.Add(richTextBox1);
            Controls.Add(btnVTdenOku);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Excel Entegrasyon";
            Load += Form1_Load;
            ResumeLayout(false);
        }

        #endregion

        private Button btnVTdenOku;
        private RichTextBox richTextBox1;
        private Button btnExceldenOku;
        private RichTextBox richTextBox2;

    }
}
