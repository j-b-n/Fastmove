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
            this.OK_button.Location = new System.Drawing.Point(287, 260);
            this.OK_button.Name = "OK_button";
            this.OK_button.Size = new System.Drawing.Size(79, 30);
            this.OK_button.TabIndex = 0;
            this.OK_button.Text = "OK";
            this.OK_button.UseVisualStyleBackColor = true;
            this.OK_button.Click += new System.EventHandler(this.OK_button_Click);
            // 
            // online_label
            // 
            this.online_label.AutoSize = true;
            this.online_label.Location = new System.Drawing.Point(12, 227);
            this.online_label.Name = "online_label";
            this.online_label.Size = new System.Drawing.Size(112, 20);
            this.online_label.TabIndex = 1;
            this.online_label.Text = "Online version:";
            // 
            // this_label
            // 
            this.this_label.AutoSize = true;
            this.this_label.Location = new System.Drawing.Point(12, 207);
            this.this_label.Name = "this_label";
            this.this_label.Size = new System.Drawing.Size(96, 20);
            this.this_label.TabIndex = 2;
            this.this_label.Text = "This version:";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(353, 145);
            this.label1.TabIndex = 3;
            this.label1.Text = "With this Outlook addin you can defer emails and move mails between folders. It i" +
    "s intended to simplify some everyday task.";
            // 
            // About
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(381, 302);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.this_label);
            this.Controls.Add(this.online_label);
            this.Controls.Add(this.OK_button);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
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