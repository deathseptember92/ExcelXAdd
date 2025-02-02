﻿namespace XAdd
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
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.Location = new System.Drawing.Point(12, 12);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(433, 412);
            this.treeView1.TabIndex = 2;
            this.treeView1.DoubleClick += new System.EventHandler(this.TreeView1_DoubleClick);
            // 
            // treeView2
            // 
            this.treeView2.Location = new System.Drawing.Point(592, 12);
            this.treeView2.Name = "treeView2";
            this.treeView2.Size = new System.Drawing.Size(433, 412);
            this.treeView2.TabIndex = 3;
            this.treeView2.DoubleClick += new System.EventHandler(this.TreeView2_DoubleClick);
            // 
            // SelectedNodesToFinal
            // 
            this.SelectedNodesToFinal.Location = new System.Drawing.Point(451, 175);
            this.SelectedNodesToFinal.Name = "SelectedNodesToFinal";
            this.SelectedNodesToFinal.Size = new System.Drawing.Size(135, 37);
            this.SelectedNodesToFinal.TabIndex = 4;
            this.SelectedNodesToFinal.Text = "=>";
            this.SelectedNodesToFinal.UseVisualStyleBackColor = true;
            this.SelectedNodesToFinal.Click += new System.EventHandler(this.SelectedNodesToFinal_Click);
            // 
            // RemoveNodesFromFinal
            // 
            this.RemoveNodesFromFinal.Location = new System.Drawing.Point(451, 218);
            this.RemoveNodesFromFinal.Name = "RemoveNodesFromFinal";
            this.RemoveNodesFromFinal.Size = new System.Drawing.Size(135, 37);
            this.RemoveNodesFromFinal.TabIndex = 5;
            this.RemoveNodesFromFinal.Text = "<=";
            this.RemoveNodesFromFinal.UseVisualStyleBackColor = true;
            this.RemoveNodesFromFinal.Click += new System.EventHandler(this.RemoveNodesFromFinal_Click);
            // 
            // AppendSheetsOK
            // 
            this.AppendSheetsOK.Location = new System.Drawing.Point(749, 443);
            this.AppendSheetsOK.Name = "AppendSheetsOK";
            this.AppendSheetsOK.Size = new System.Drawing.Size(135, 37);
            this.AppendSheetsOK.TabIndex = 6;
            this.AppendSheetsOK.Text = "ОК";
            this.AppendSheetsOK.UseVisualStyleBackColor = true;
            this.AppendSheetsOK.Click += new System.EventHandler(this.AppendSheetsOK_Click);
            // 
            // AppendSheetsCancel
            // 
            this.AppendSheetsCancel.Location = new System.Drawing.Point(890, 443);
            this.AppendSheetsCancel.Name = "AppendSheetsCancel";
            this.AppendSheetsCancel.Size = new System.Drawing.Size(135, 37);
            this.AppendSheetsCancel.TabIndex = 7;
            this.AppendSheetsCancel.Text = "Отмена";
            this.AppendSheetsCancel.UseVisualStyleBackColor = true;
            this.AppendSheetsCancel.Click += new System.EventHandler(this.AppendSheetsCancel_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(13, 443);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(287, 17);
            this.checkBox1.TabIndex = 8;
            this.checkBox1.Text = "У листов одинаковые поля. Свести в одну таблицу.";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // AppendSheetsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1037, 492);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.AppendSheetsCancel);
            this.Controls.Add(this.AppendSheetsOK);
            this.Controls.Add(this.RemoveNodesFromFinal);
            this.Controls.Add(this.SelectedNodesToFinal);
            this.Controls.Add(this.treeView2);
            this.Controls.Add(this.treeView1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "AppendSheetsForm";
            this.Text = "Объединение листов";
            this.Deactivate += new System.EventHandler(this.AppendSheetsForm_Deactivate);
            this.Load += new System.EventHandler(this.AppendSheetsForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        internal System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.Button SelectedNodesToFinal;
        private System.Windows.Forms.Button RemoveNodesFromFinal;
        internal System.Windows.Forms.TreeView treeView2;
        private System.Windows.Forms.Button AppendSheetsOK;
        private System.Windows.Forms.Button AppendSheetsCancel;
        internal System.Windows.Forms.CheckBox checkBox1;
    }
}