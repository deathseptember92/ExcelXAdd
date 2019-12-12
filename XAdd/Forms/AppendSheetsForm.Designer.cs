namespace XAdd
{
    partial class AppendSheetsForm
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
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.treeView2 = new System.Windows.Forms.TreeView();
            this.SelectedNodesToFinal = new System.Windows.Forms.Button();
            this.RemoveNodesFromFinal = new System.Windows.Forms.Button();
            this.AppendSheetsOK = new System.Windows.Forms.Button();
            this.AppendSheetsCancel = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.Location = new System.Drawing.Point(6, 19);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(367, 399);
            this.treeView1.TabIndex = 2;
            // 
            // treeView2
            // 
            this.treeView2.Location = new System.Drawing.Point(6, 19);
            this.treeView2.Name = "treeView2";
            this.treeView2.Size = new System.Drawing.Size(367, 399);
            this.treeView2.TabIndex = 3;
            // 
            // SelectedNodesToFinal
            // 
            this.SelectedNodesToFinal.BackColor = System.Drawing.Color.WhiteSmoke;
            this.SelectedNodesToFinal.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SelectedNodesToFinal.Location = new System.Drawing.Point(400, 175);
            this.SelectedNodesToFinal.Name = "SelectedNodesToFinal";
            this.SelectedNodesToFinal.Size = new System.Drawing.Size(187, 37);
            this.SelectedNodesToFinal.TabIndex = 4;
            this.SelectedNodesToFinal.Text = "=>";
            this.SelectedNodesToFinal.UseVisualStyleBackColor = false;
            // 
            // RemoveNodesFromFinal
            // 
            this.RemoveNodesFromFinal.BackColor = System.Drawing.Color.WhiteSmoke;
            this.RemoveNodesFromFinal.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.RemoveNodesFromFinal.Location = new System.Drawing.Point(400, 218);
            this.RemoveNodesFromFinal.Name = "RemoveNodesFromFinal";
            this.RemoveNodesFromFinal.Size = new System.Drawing.Size(187, 37);
            this.RemoveNodesFromFinal.TabIndex = 5;
            this.RemoveNodesFromFinal.Text = "<=";
            this.RemoveNodesFromFinal.UseVisualStyleBackColor = false;
            // 
            // AppendSheetsOK
            // 
            this.AppendSheetsOK.BackColor = System.Drawing.Color.WhiteSmoke;
            this.AppendSheetsOK.Location = new System.Drawing.Point(694, 443);
            this.AppendSheetsOK.Name = "AppendSheetsOK";
            this.AppendSheetsOK.Size = new System.Drawing.Size(135, 37);
            this.AppendSheetsOK.TabIndex = 6;
            this.AppendSheetsOK.Text = "ОК";
            this.AppendSheetsOK.UseVisualStyleBackColor = false;
            // 
            // AppendSheetsCancel
            // 
            this.AppendSheetsCancel.BackColor = System.Drawing.Color.WhiteSmoke;
            this.AppendSheetsCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.AppendSheetsCancel.Location = new System.Drawing.Point(835, 443);
            this.AppendSheetsCancel.Name = "AppendSheetsCancel";
            this.AppendSheetsCancel.Size = new System.Drawing.Size(135, 37);
            this.AppendSheetsCancel.TabIndex = 7;
            this.AppendSheetsCancel.Text = "Отмена";
            this.AppendSheetsCancel.UseVisualStyleBackColor = false;
            this.AppendSheetsCancel.Click += new System.EventHandler(this.AppendSheetsCancel_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(6, 64);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(164, 17);
            this.checkBox1.TabIndex = 8;
            this.checkBox1.Text = "У листов одинаковые поля";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.treeView1);
            this.groupBox1.Location = new System.Drawing.Point(13, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(381, 424);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Список листов";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.treeView2);
            this.groupBox2.Location = new System.Drawing.Point(593, 13);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(377, 424);
            this.groupBox2.TabIndex = 10;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Добавленные листы/книги";
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(19, 443);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(170, 17);
            this.checkBox2.TabIndex = 11;
            this.checkBox2.Text = "Отображать скрытые листы";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Location = new System.Drawing.Point(6, 19);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(129, 17);
            this.checkBox3.TabIndex = 12;
            this.checkBox3.Text = "Учитывать фильтры";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.Location = new System.Drawing.Point(6, 41);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(144, 17);
            this.checkBox4.TabIndex = 13;
            this.checkBox4.Text = "Калькуляция включена";
            this.checkBox4.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.checkBox4);
            this.groupBox3.Controls.Add(this.checkBox3);
            this.groupBox3.Controls.Add(this.checkBox1);
            this.groupBox3.Location = new System.Drawing.Point(400, 344);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(187, 87);
            this.groupBox3.TabIndex = 14;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Настройки объединения";
            // 
            // AppendSheetsForm
            // 
            this.AcceptButton = this.AppendSheetsOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.CancelButton = this.AppendSheetsCancel;
            this.ClientSize = new System.Drawing.Size(982, 488);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.checkBox2);
            this.Controls.Add(this.AppendSheetsCancel);
            this.Controls.Add(this.AppendSheetsOK);
            this.Controls.Add(this.RemoveNodesFromFinal);
            this.Controls.Add(this.SelectedNodesToFinal);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AppendSheetsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Объединение листов - XAdd";
            this.Deactivate += new System.EventHandler(this.AppendSheetsForm_Deactivate);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.AppendSheetsForm_FormClosing);
            this.Load += new System.EventHandler(this.AppendSheetsForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        internal System.Windows.Forms.TreeView treeView1;
        internal System.Windows.Forms.TreeView treeView2;
        internal System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        internal System.Windows.Forms.CheckBox checkBox2;
        internal System.Windows.Forms.Button SelectedNodesToFinal;
        internal System.Windows.Forms.Button RemoveNodesFromFinal;
        internal System.Windows.Forms.Button AppendSheetsOK;
        internal System.Windows.Forms.Button AppendSheetsCancel;
        internal System.Windows.Forms.CheckBox checkBox3;
        private System.Windows.Forms.GroupBox groupBox3;
        internal System.Windows.Forms.CheckBox checkBox4;
    }
}