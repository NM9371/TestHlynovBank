namespace TestHlynovBank
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
            this.EnterButton = new System.Windows.Forms.Button();
            this.ComboBox1 = new System.Windows.Forms.ComboBox();
            this.grid = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
            this.SuspendLayout();
            // 
            // EnterButton
            // 
            this.EnterButton.Location = new System.Drawing.Point(120, 283);
            this.EnterButton.Name = "EnterButton";
            this.EnterButton.Size = new System.Drawing.Size(75, 23);
            this.EnterButton.TabIndex = 3;
            this.EnterButton.Text = "OK";
            this.EnterButton.UseVisualStyleBackColor = true;
            this.EnterButton.Visible = false;
            this.EnterButton.Click += new System.EventHandler(this.EnterButton_Click);
            // 
            // ComboBox1
            // 
            this.ComboBox1.DisplayMember = "fileEntries";
            this.ComboBox1.FormattingEnabled = true;
            this.ComboBox1.Location = new System.Drawing.Point(154, 36);
            this.ComboBox1.Name = "ComboBox1";
            this.ComboBox1.Size = new System.Drawing.Size(121, 21);
            this.ComboBox1.TabIndex = 14;
            this.ComboBox1.ValueMember = "fileEntries";
            this.ComboBox1.SelectedIndexChanged += new System.EventHandler(this.ComboBox1_SelectedIndexChanged);
            // 
            // grid
            // 
            this.grid.AllowUserToAddRows = false;
            this.grid.AllowUserToDeleteRows = false;
            this.grid.AllowUserToResizeRows = false;
            this.grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.grid.Location = new System.Drawing.Point(10, 74);
            this.grid.Name = "grid";
            this.grid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.grid.RowHeadersVisible = false;
            this.grid.Size = new System.Drawing.Size(304, 186);
            this.grid.TabIndex = 17;
            this.grid.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(38, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(104, 13);
            this.label1.TabIndex = 18;
            this.label1.Text = "Название шаблона";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(324, 327);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.grid);
            this.Controls.Add(this.ComboBox1);
            this.Controls.Add(this.EnterButton);
            this.Name = "Form1";
            this.Text = "Заполнение шаблона";
            ((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button EnterButton;
        private System.Windows.Forms.ComboBox ComboBox1;
        private System.Windows.Forms.DataGridView grid;
        private System.Windows.Forms.Label label1;
    }
}

