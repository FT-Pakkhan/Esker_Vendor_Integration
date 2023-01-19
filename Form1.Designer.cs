
namespace Esker_Vendor_Integration
{
    partial class FTS00OVI
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
            this.Log = new System.Windows.Forms.RichTextBox();
            this.btnExecute = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Log
            // 
            this.Log.Location = new System.Drawing.Point(25, 12);
            this.Log.Name = "Log";
            this.Log.Size = new System.Drawing.Size(745, 189);
            this.Log.TabIndex = 0;
            this.Log.Text = "";
            // 
            // btnExecute
            // 
            this.btnExecute.Location = new System.Drawing.Point(25, 221);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(745, 23);
            this.btnExecute.TabIndex = 1;
            this.btnExecute.Text = "Execute";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // FTS00OVI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 275);
            this.Controls.Add(this.btnExecute);
            this.Controls.Add(this.Log);
            this.Name = "FTS00OVI";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Esker Vendor Integration";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FTS00OVI_FormClosing);
            this.Load += new System.EventHandler(this.FTS00OVI_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btnExecute;
        public System.Windows.Forms.RichTextBox Log;
    }
}

