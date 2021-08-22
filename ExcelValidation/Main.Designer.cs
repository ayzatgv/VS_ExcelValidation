
namespace ExcelValidation
{
    partial class Main
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
            this.Button_Import = new System.Windows.Forms.Button();
            this.Button_Validate = new System.Windows.Forms.Button();
            this.Button_Create = new System.Windows.Forms.Button();
            this.Button_Clear = new System.Windows.Forms.Button();
            this.RichText = new System.Windows.Forms.RichTextBox();
            this.Label_ErrorCount = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Button_Import
            // 
            this.Button_Import.Location = new System.Drawing.Point(13, 13);
            this.Button_Import.Name = "Button_Import";
            this.Button_Import.Size = new System.Drawing.Size(75, 23);
            this.Button_Import.TabIndex = 0;
            this.Button_Import.Text = "Import";
            this.Button_Import.UseVisualStyleBackColor = true;
            this.Button_Import.Click += new System.EventHandler(this.Button_Import_Click);
            // 
            // Button_Validate
            // 
            this.Button_Validate.Location = new System.Drawing.Point(94, 13);
            this.Button_Validate.Name = "Button_Validate";
            this.Button_Validate.Size = new System.Drawing.Size(75, 23);
            this.Button_Validate.TabIndex = 1;
            this.Button_Validate.Text = "Validate";
            this.Button_Validate.UseVisualStyleBackColor = true;
            this.Button_Validate.Click += new System.EventHandler(this.Button_Validate_Click);
            // 
            // Button_Create
            // 
            this.Button_Create.Location = new System.Drawing.Point(175, 13);
            this.Button_Create.Name = "Button_Create";
            this.Button_Create.Size = new System.Drawing.Size(75, 23);
            this.Button_Create.TabIndex = 2;
            this.Button_Create.Text = "Create";
            this.Button_Create.UseVisualStyleBackColor = true;
            this.Button_Create.Click += new System.EventHandler(this.Button_Create_Click);
            // 
            // Button_Clear
            // 
            this.Button_Clear.Location = new System.Drawing.Point(256, 13);
            this.Button_Clear.Name = "Button_Clear";
            this.Button_Clear.Size = new System.Drawing.Size(75, 23);
            this.Button_Clear.TabIndex = 3;
            this.Button_Clear.Text = "Clear";
            this.Button_Clear.UseVisualStyleBackColor = true;
            this.Button_Clear.Click += new System.EventHandler(this.Button_Clear_Click);
            // 
            // RichText
            // 
            this.RichText.Location = new System.Drawing.Point(13, 42);
            this.RichText.Name = "RichText";
            this.RichText.ReadOnly = true;
            this.RichText.Size = new System.Drawing.Size(759, 307);
            this.RichText.TabIndex = 4;
            this.RichText.Text = "";
            // 
            // Label_ErrorCount
            // 
            this.Label_ErrorCount.AutoSize = true;
            this.Label_ErrorCount.Location = new System.Drawing.Point(337, 17);
            this.Label_ErrorCount.Name = "Label_ErrorCount";
            this.Label_ErrorCount.Size = new System.Drawing.Size(13, 15);
            this.Label_ErrorCount.TabIndex = 5;
            this.Label_ErrorCount.Text = "0";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 361);
            this.Controls.Add(this.Label_ErrorCount);
            this.Controls.Add(this.RichText);
            this.Controls.Add(this.Button_Clear);
            this.Controls.Add(this.Button_Create);
            this.Controls.Add(this.Button_Validate);
            this.Controls.Add(this.Button_Import);
            this.Name = "Main";
            this.Text = "Main";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Main_FormClosing);
            this.Load += new System.EventHandler(this.Main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Button_Import;
        private System.Windows.Forms.Button Button_Validate;
        private System.Windows.Forms.Button Button_Create;
        private System.Windows.Forms.Button Button_Clear;
        private System.Windows.Forms.RichTextBox RichText;
        private System.Windows.Forms.Label Label_ErrorCount;
    }
}