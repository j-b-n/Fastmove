namespace FastMove
{
    partial class About
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(About));
            this.OK_button = new System.Windows.Forms.Button();
            this.online_label = new System.Windows.Forms.Label();
            this.this_label = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // OK_button
            // 
            this.OK_button.Location = new System.Drawing.Point(191, 169);
            this.OK_button.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.OK_button.Name = "OK_button";
            this.OK_button.Size = new System.Drawing.Size(53, 19);
            this.OK_button.TabIndex = 0;
            this.OK_button.Text = "OK";
            this.OK_button.UseVisualStyleBackColor = true;
            this.OK_button.Click += new System.EventHandler(this.OK_button_Click);
            // 
            // online_label
            // 
            this.online_label.AutoSize = true;
            this.online_label.Location = new System.Drawing.Point(8, 148);
            this.online_label.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.online_label.Name = "online_label";
            this.online_label.Size = new System.Drawing.Size(77, 13);
            this.online_label.TabIndex = 1;
            this.online_label.Text = "Online version:";
            // 
            // this_label
            // 
            this.this_label.AutoSize = true;
            this.this_label.Location = new System.Drawing.Point(8, 124);
            this.this_label.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.this_label.Name = "this_label";
            this.this_label.Size = new System.Drawing.Size(67, 13);
            this.this_label.TabIndex = 2;
            this.this_label.Text = "This version:";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(9, 8);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(235, 94);
            this.label1.TabIndex = 3;
            this.label1.Text = "With this Outlook addin you can defer emails and move mails between folders. It i" +
    "s intended to simplify some everyday task.";
            // 
            // About
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(254, 196);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.this_label);
            this.Controls.Add(this.online_label);
            this.Controls.Add(this.OK_button);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "About";
            this.Text = "About";
            this.Load += new System.EventHandler(this.About_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button OK_button;
        private System.Windows.Forms.Label online_label;
        private System.Windows.Forms.Label this_label;
        private System.Windows.Forms.Label label1;
    }
}