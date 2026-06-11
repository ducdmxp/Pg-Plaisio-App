namespace Convert2DTo3D.UI
{
    partial class PreviewFrm
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.elemHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.elemHost2 = new System.Windows.Forms.Integration.ElementHost();
            this.button1 = new System.Windows.Forms.Button();
            this.picPreview1 = new System.Windows.Forms.PictureBox();
            this.picPreview2 = new System.Windows.Forms.PictureBox();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picPreview1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picPreview2)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.elemHost1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.elemHost2, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.button1, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.picPreview1, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.picPreview2, 1, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 84F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(641, 622);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // elemHost1
            // 
            this.elemHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elemHost1.Location = new System.Drawing.Point(3, 3);
            this.elemHost1.Name = "elemHost1";
            this.elemHost1.Size = new System.Drawing.Size(314, 263);
            this.elemHost1.TabIndex = 0;
            this.elemHost1.Text = "elementHost1";
            this.elemHost1.Child = null;
            // 
            // elemHost2
            // 
            this.elemHost2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elemHost2.Location = new System.Drawing.Point(323, 3);
            this.elemHost2.Name = "elemHost2";
            this.elemHost2.Size = new System.Drawing.Size(315, 263);
            this.elemHost2.TabIndex = 1;
            this.elemHost2.Text = "elementHost1";
            this.elemHost2.Child = null;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(3, 541);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(125, 42);
            this.button1.TabIndex = 2;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // picPreview1
            // 
            this.picPreview1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.picPreview1.Location = new System.Drawing.Point(3, 272);
            this.picPreview1.Name = "picPreview1";
            this.picPreview1.Size = new System.Drawing.Size(314, 263);
            this.picPreview1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picPreview1.TabIndex = 3;
            this.picPreview1.TabStop = false;
            // 
            // picPreview2
            // 
            this.picPreview2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.picPreview2.Location = new System.Drawing.Point(323, 272);
            this.picPreview2.Name = "picPreview2";
            this.picPreview2.Size = new System.Drawing.Size(315, 263);
            this.picPreview2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picPreview2.TabIndex = 4;
            this.picPreview2.TabStop = false;
            // 
            // PreviewFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(641, 622);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "PreviewFrm";
            this.Text = "PreviewFrm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.PreviewFrm_FormClosing);
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picPreview1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picPreview2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Integration.ElementHost elemHost1;
        private System.Windows.Forms.Integration.ElementHost elemHost2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.PictureBox picPreview1;
        private System.Windows.Forms.PictureBox picPreview2;
    }
}