﻿namespace XAdd
{
    partial class AppendWorkbooksForm
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
            this.buttonAppend = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonFileDialog = new System.Windows.Forms.Button();
            this.buttonAdd = new System.Windows.Forms.Button();
            this.buttonExclude = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonAppend
            // 
            this.buttonAppend.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonAppend.Location = new System.Drawing.Point(522, 258);
            this.buttonAppend.Name = "buttonAppend";
            this.buttonAppend.Size = new System.Drawing.Size(85, 23);
            this.buttonAppend.TabIndex = 0;
            this.buttonAppend.Text = "Объединить";
            this.buttonAppend.UseVisualStyleBackColor = true;
            // 
            // listView1
            // 
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(12, 12);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(660, 240);
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(613, 258);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(59, 23);
            this.buttonCancel.TabIndex = 2;
            this.buttonCancel.Text = "Отмена";
            this.buttonCancel.UseVisualStyleBackColor = true;
            // 
            // buttonFileDialog
            // 
            this.buttonFileDialog.Location = new System.Drawing.Point(13, 258);
            this.buttonFileDialog.Name = "buttonFileDialog";
            this.buttonFileDialog.Size = new System.Drawing.Size(53, 23);
            this.buttonFileDialog.TabIndex = 3;
            this.buttonFileDialog.Text = "...";
            this.buttonFileDialog.UseVisualStyleBackColor = true;
            // 
            // buttonAdd
            // 
            this.buttonAdd.Location = new System.Drawing.Point(72, 258);
            this.buttonAdd.Name = "buttonAdd";
            this.buttonAdd.Size = new System.Drawing.Size(73, 23);
            this.buttonAdd.TabIndex = 4;
            this.buttonAdd.Text = "Добавить";
            this.buttonAdd.UseVisualStyleBackColor = true;
            // 
            // buttonExclude
            // 
            this.buttonExclude.Location = new System.Drawing.Point(151, 258);
            this.buttonExclude.Name = "buttonExclude";
            this.buttonExclude.Size = new System.Drawing.Size(72, 23);
            this.buttonExclude.TabIndex = 5;
            this.buttonExclude.Text = "Исключить";
            this.buttonExclude.UseVisualStyleBackColor = true;
            // 
            // AppendWorkbooksForm
            // 
            this.AcceptButton = this.buttonAppend;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.CancelButton = this.buttonCancel;
            this.ClientSize = new System.Drawing.Size(684, 288);
            this.Controls.Add(this.buttonExclude);
            this.Controls.Add(this.buttonAdd);
            this.Controls.Add(this.buttonFileDialog);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.buttonAppend);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Name = "AppendWorkbooksForm";
            this.Text = "Объединение книг - XAdd";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.AppendWorkbooksForm_FormClosing);
            this.Load += new System.EventHandler(this.AppendWorkbooksForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonAppend;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonFileDialog;
        private System.Windows.Forms.Button buttonAdd;
        private System.Windows.Forms.Button buttonExclude;
    }
}