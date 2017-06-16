namespace DrawBallot
{
    partial class DataForm
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
            this.loadButton = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.bAdd = new System.Windows.Forms.Button();
            this.bDel = new System.Windows.Forms.Button();
            this.bMod = new System.Windows.Forms.Button();
            this.IDBox = new System.Windows.Forms.TextBox();
            this.FNBox = new System.Windows.Forms.TextBox();
            this.LNBox = new System.Windows.Forms.TextBox();
            this.bCon = new System.Windows.Forms.Button();
            this.bCan = new System.Windows.Forms.Button();
            this.bReset = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // loadButton
            // 
            this.loadButton.Location = new System.Drawing.Point(12, 12);
            this.loadButton.Name = "loadButton";
            this.loadButton.Size = new System.Drawing.Size(75, 23);
            this.loadButton.TabIndex = 0;
            this.loadButton.Text = "Load file";
            this.loadButton.UseVisualStyleBackColor = true;
            this.loadButton.Click += new System.EventHandler(this.loadButton_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 99);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(627, 150);
            this.dataGridView1.TabIndex = 1;
            // 
            // bAdd
            // 
            this.bAdd.Location = new System.Drawing.Point(93, 12);
            this.bAdd.Name = "bAdd";
            this.bAdd.Size = new System.Drawing.Size(75, 23);
            this.bAdd.TabIndex = 2;
            this.bAdd.Text = "Add Row";
            this.bAdd.UseVisualStyleBackColor = true;
            this.bAdd.Click += new System.EventHandler(this.bAdd_Click);
            // 
            // bDel
            // 
            this.bDel.Location = new System.Drawing.Point(174, 12);
            this.bDel.Name = "bDel";
            this.bDel.Size = new System.Drawing.Size(75, 23);
            this.bDel.TabIndex = 2;
            this.bDel.Text = "Delete Row";
            this.bDel.UseVisualStyleBackColor = true;
            this.bDel.Click += new System.EventHandler(this.bDel_Click);
            // 
            // bMod
            // 
            this.bMod.Location = new System.Drawing.Point(255, 12);
            this.bMod.Name = "bMod";
            this.bMod.Size = new System.Drawing.Size(75, 23);
            this.bMod.TabIndex = 2;
            this.bMod.Text = "Modify Row";
            this.bMod.UseVisualStyleBackColor = true;
            this.bMod.Click += new System.EventHandler(this.bMod_Click);
            // 
            // boxID
            // 
            this.IDBox.Location = new System.Drawing.Point(12, 41);
            this.IDBox.Name = "boxID";
            this.IDBox.Size = new System.Drawing.Size(75, 20);
            this.IDBox.TabIndex = 3;
            this.IDBox.Visible = false;
            // 
            // boxFN
            // 
            this.FNBox.Location = new System.Drawing.Point(93, 41);
            this.FNBox.Name = "boxFN";
            this.FNBox.Size = new System.Drawing.Size(75, 20);
            this.FNBox.TabIndex = 3;
            this.FNBox.Visible = false;
            // 
            // boxLN
            // 
            this.LNBox.Location = new System.Drawing.Point(174, 41);
            this.LNBox.Name = "boxLN";
            this.LNBox.Size = new System.Drawing.Size(75, 20);
            this.LNBox.TabIndex = 3;
            this.LNBox.Visible = false;
            // 
            // bCon
            // 
            this.bCon.Location = new System.Drawing.Point(12, 70);
            this.bCon.Name = "bCon";
            this.bCon.Size = new System.Drawing.Size(75, 23);
            this.bCon.TabIndex = 4;
            this.bCon.Text = "Confirm";
            this.bCon.UseVisualStyleBackColor = true;
            this.bCon.Visible = false;
            this.bCon.Click += new System.EventHandler(this.bCon_Click);
            // 
            // bCan
            // 
            this.bCan.Location = new System.Drawing.Point(93, 70);
            this.bCan.Name = "bCan";
            this.bCan.Size = new System.Drawing.Size(75, 23);
            this.bCan.TabIndex = 4;
            this.bCan.Text = "Cancel";
            this.bCan.UseVisualStyleBackColor = true;
            this.bCan.Visible = false;
            this.bCan.Click += new System.EventHandler(this.bCan_Click);
            // 
            // bReset
            // 
            this.bReset.Location = new System.Drawing.Point(336, 12);
            this.bReset.Name = "bReset";
            this.bReset.Size = new System.Drawing.Size(75, 23);
            this.bReset.TabIndex = 5;
            this.bReset.Text = "Reset";
            this.bReset.UseVisualStyleBackColor = true;
            this.bReset.Click += new System.EventHandler(this.bReset_Click);
            // 
            // DataForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(651, 261);
            this.Controls.Add(this.bReset);
            this.Controls.Add(this.bCan);
            this.Controls.Add(this.bCon);
            this.Controls.Add(this.LNBox);
            this.Controls.Add(this.FNBox);
            this.Controls.Add(this.IDBox);
            this.Controls.Add(this.bMod);
            this.Controls.Add(this.bDel);
            this.Controls.Add(this.bAdd);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.loadButton);
            this.Name = "DataForm";
            this.Text = "DataForm";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.DataForm_FormClosed);
            this.Load += new System.EventHandler(this.DataForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button loadButton;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button bAdd;
        private System.Windows.Forms.Button bDel;
        private System.Windows.Forms.Button bMod;
        private System.Windows.Forms.TextBox IDBox;
        private System.Windows.Forms.TextBox FNBox;
        private System.Windows.Forms.TextBox LNBox;
        private System.Windows.Forms.Button bCon;
        private System.Windows.Forms.Button bCan;
        private System.Windows.Forms.Button bReset;
    }
}