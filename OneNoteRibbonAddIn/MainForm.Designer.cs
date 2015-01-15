namespace OneNoteRibbonAddIn
{
    partial class MainForm
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
            this.btnEnumerateNotebooks = new System.Windows.Forms.Button();
            this.btnEnumerateSections = new System.Windows.Forms.Button();
            this.btnGetPageTitle = new System.Windows.Forms.Button();
            this.btnGetPageContent = new System.Windows.Forms.Button();
            this.btnUpdatePageContent = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnEnumerateNotebooks
            // 
            this.btnEnumerateNotebooks.Location = new System.Drawing.Point(13, 13);
            this.btnEnumerateNotebooks.Name = "btnEnumerateNotebooks";
            this.btnEnumerateNotebooks.Size = new System.Drawing.Size(149, 23);
            this.btnEnumerateNotebooks.TabIndex = 0;
            this.btnEnumerateNotebooks.Text = "Enumerate Notebooks";
            this.btnEnumerateNotebooks.UseVisualStyleBackColor = true;
            this.btnEnumerateNotebooks.Click += new System.EventHandler(this.btnEnumerateNotebooks_Click);
            // 
            // btnEnumerateSections
            // 
            this.btnEnumerateSections.Location = new System.Drawing.Point(13, 42);
            this.btnEnumerateSections.Name = "btnEnumerateSections";
            this.btnEnumerateSections.Size = new System.Drawing.Size(149, 23);
            this.btnEnumerateSections.TabIndex = 1;
            this.btnEnumerateSections.Text = "Enumerate Sections";
            this.btnEnumerateSections.UseVisualStyleBackColor = true;
            this.btnEnumerateSections.Click += new System.EventHandler(this.btnEnumerateSections_Click);
            // 
            // btnGetPageTitle
            // 
            this.btnGetPageTitle.Location = new System.Drawing.Point(12, 71);
            this.btnGetPageTitle.Name = "btnGetPageTitle";
            this.btnGetPageTitle.Size = new System.Drawing.Size(150, 23);
            this.btnGetPageTitle.TabIndex = 2;
            this.btnGetPageTitle.Text = "Get Page Title";
            this.btnGetPageTitle.UseVisualStyleBackColor = true;
            this.btnGetPageTitle.Click += new System.EventHandler(this.btnGetPageTitle_Click);
            // 
            // btnGetPageContent
            // 
            this.btnGetPageContent.Location = new System.Drawing.Point(12, 100);
            this.btnGetPageContent.Name = "btnGetPageContent";
            this.btnGetPageContent.Size = new System.Drawing.Size(150, 23);
            this.btnGetPageContent.TabIndex = 3;
            this.btnGetPageContent.Text = "Get Page Content";
            this.btnGetPageContent.UseVisualStyleBackColor = true;
            this.btnGetPageContent.Click += new System.EventHandler(this.btnGetPageContent_Click);
            // 
            // btnUpdatePageContent
            // 
            this.btnUpdatePageContent.Location = new System.Drawing.Point(13, 129);
            this.btnUpdatePageContent.Name = "btnUpdatePageContent";
            this.btnUpdatePageContent.Size = new System.Drawing.Size(149, 23);
            this.btnUpdatePageContent.TabIndex = 4;
            this.btnUpdatePageContent.Text = "Update Page Content";
            this.btnUpdatePageContent.UseVisualStyleBackColor = true;
            this.btnUpdatePageContent.Click += new System.EventHandler(this.btnUpdatePageContent_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(174, 164);
            this.Controls.Add(this.btnUpdatePageContent);
            this.Controls.Add(this.btnGetPageContent);
            this.Controls.Add(this.btnGetPageTitle);
            this.Controls.Add(this.btnEnumerateSections);
            this.Controls.Add(this.btnEnumerateNotebooks);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.Text = "Tools";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnEnumerateNotebooks;
        private System.Windows.Forms.Button btnEnumerateSections;
        private System.Windows.Forms.Button btnGetPageTitle;
        private System.Windows.Forms.Button btnGetPageContent;
        private System.Windows.Forms.Button btnUpdatePageContent;
    }
}