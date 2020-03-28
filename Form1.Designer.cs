namespace FunWithExcel
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.load_sheet_excel = new System.Windows.Forms.Button();
            this.info = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // load_sheet_excel
            // 
            this.load_sheet_excel.Location = new System.Drawing.Point(12, 13);
            this.load_sheet_excel.Name = "load_sheet_excel";
            this.load_sheet_excel.Size = new System.Drawing.Size(160, 23);
            this.load_sheet_excel.TabIndex = 0;
            this.load_sheet_excel.Text = "Load";
            this.load_sheet_excel.UseVisualStyleBackColor = true;
            this.load_sheet_excel.Click += new System.EventHandler(this.load_sheet_excel_Click);
            // 
            // info
            // 
            this.info.Location = new System.Drawing.Point(13, 60);
            this.info.Name = "info";
            this.info.Size = new System.Drawing.Size(558, 275);
            this.info.TabIndex = 1;
            this.info.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(583, 347);
            this.Controls.Add(this.info);
            this.Controls.Add(this.load_sheet_excel);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button load_sheet_excel;
        private System.Windows.Forms.RichTextBox info;
    }
}

