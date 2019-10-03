namespace XAdd
{
    partial class SheetsManagerForm
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
            this.components = new System.ComponentModel.Container();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.NewBookButton = new System.Windows.Forms.Button();
            this.NewSheetButton = new System.Windows.Forms.Button();
            this.OpenButton = new System.Windows.Forms.Button();
            this.RemoveButton = new System.Windows.Forms.Button();
            this.RenameButton = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.CheckBoxes = true;
            this.treeView1.HideSelection = false;
            this.treeView1.Location = new System.Drawing.Point(18, 31);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(304, 411);
            this.treeView1.TabIndex = 0;
            this.treeView1.BeforeCheck += new System.Windows.Forms.TreeViewCancelEventHandler(this.TreeView1_BeforeCheck);
            this.treeView1.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.TreeView1_NodeMouseClick);
            this.treeView1.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.TreeView1_NodeMouseDoubleClick);
            this.treeView1.MouseEnter += new System.EventHandler(this.TreeView1_MouseEnter);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.NewBookButton);
            this.groupBox1.Controls.Add(this.NewSheetButton);
            this.groupBox1.Controls.Add(this.OpenButton);
            this.groupBox1.Controls.Add(this.RemoveButton);
            this.groupBox1.Controls.Add(this.RenameButton);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(316, 480);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Список листов";
            // 
            // NewBookButton
            // 
            this.NewBookButton.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.NewBookButton.BackgroundImage = global::XAdd.Properties.Resources.newdocument;
            this.NewBookButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.NewBookButton.Location = new System.Drawing.Point(52, 434);
            this.NewBookButton.Name = "NewBookButton";
            this.NewBookButton.Size = new System.Drawing.Size(40, 40);
            this.NewBookButton.TabIndex = 4;
            this.NewBookButton.UseVisualStyleBackColor = false;
            this.NewBookButton.Click += new System.EventHandler(this.NewBookButton_Click);
            // 
            // NewSheetButton
            // 
            this.NewSheetButton.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.NewSheetButton.BackgroundImage = global::XAdd.Properties.Resources.newsheet;
            this.NewSheetButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.NewSheetButton.Location = new System.Drawing.Point(178, 434);
            this.NewSheetButton.Name = "NewSheetButton";
            this.NewSheetButton.Size = new System.Drawing.Size(40, 40);
            this.NewSheetButton.TabIndex = 3;
            this.NewSheetButton.UseVisualStyleBackColor = false;
            this.NewSheetButton.Click += new System.EventHandler(this.NewSheetButton_Click);
            // 
            // OpenButton
            // 
            this.OpenButton.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.OpenButton.BackgroundImage = global::XAdd.Properties.Resources.open1;
            this.OpenButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.OpenButton.Location = new System.Drawing.Point(6, 434);
            this.OpenButton.Name = "OpenButton";
            this.OpenButton.Size = new System.Drawing.Size(40, 40);
            this.OpenButton.TabIndex = 2;
            this.OpenButton.UseVisualStyleBackColor = false;
            this.OpenButton.Click += new System.EventHandler(this.OpenButton_Click);
            // 
            // RemoveButton
            // 
            this.RemoveButton.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.RemoveButton.BackgroundImage = global::XAdd.Properties.Resources._159_1597907_delete_garbage_remove_trash_trash_can_icon_delete1;
            this.RemoveButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.RemoveButton.Location = new System.Drawing.Point(270, 434);
            this.RemoveButton.Name = "RemoveButton";
            this.RemoveButton.Size = new System.Drawing.Size(40, 40);
            this.RemoveButton.TabIndex = 1;
            this.RemoveButton.UseVisualStyleBackColor = false;
            this.RemoveButton.Click += new System.EventHandler(this.RemoveButton_Click);
            // 
            // RenameButton
            // 
            this.RenameButton.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.RenameButton.BackgroundImage = global::XAdd.Properties.Resources.rename;
            this.RenameButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.RenameButton.Location = new System.Drawing.Point(224, 434);
            this.RenameButton.Name = "RenameButton";
            this.RenameButton.Size = new System.Drawing.Size(40, 40);
            this.RenameButton.TabIndex = 0;
            this.RenameButton.UseVisualStyleBackColor = false;
            this.RenameButton.Click += new System.EventHandler(this.RenameButton_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.panel1);
            this.groupBox2.Location = new System.Drawing.Point(334, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(641, 480);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Предпросмотр";
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Location = new System.Drawing.Point(7, 20);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(628, 454);
            this.panel1.TabIndex = 0;
            this.panel1.MouseEnter += new System.EventHandler(this.Panel1_MouseEnter);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(3, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(564, 407);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.MouseEnter += new System.EventHandler(this.PictureBox1_MouseEnter);
            // 
            // SheetsManagerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(982, 504);
            this.Controls.Add(this.treeView1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SheetsManagerForm";
            this.Text = "Диспетчер листов - XAdd";
            this.Activated += new System.EventHandler(this.SheetsManagerForm_Activated);
            this.Deactivate += new System.EventHandler(this.SheetsManagerForm_Deactivate);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SheetsManagerForm_FormClosing);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Panel panel1;
        internal System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button RenameButton;
        private System.Windows.Forms.Button OpenButton;
        private System.Windows.Forms.Button RemoveButton;
        private System.Windows.Forms.Button NewBookButton;
        private System.Windows.Forms.Button NewSheetButton;
    }
}