﻿namespace XAdd
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.FormulaFormat = this.Factory.CreateRibbonToggleButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.RemoveColumns = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.AppendSheets = this.Factory.CreateRibbonButton();
            this.AppendSheetsCustom = this.Factory.CreateRibbonButton();
            this.TableOfContents = this.Factory.CreateRibbonButton();
            this.SheetsManager = this.Factory.CreateRibbonButton();
            this.toggleButton1 = this.Factory.CreateRibbonToggleButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.InsertDate = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.Currency = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group5.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "XAdd";
            this.tab1.Name = "tab1";
            // 
            // group5
            // 
            this.group5.Items.Add(this.FormulaFormat);
            this.group5.Label = "Общее";
            this.group5.Name = "group5";
            // 
            // FormulaFormat
            // 
            this.FormulaFormat.Label = "Стиль ссылок R1C1";
            this.FormulaFormat.Name = "FormulaFormat";
            this.FormulaFormat.OfficeImageId = "WordCountList";
            this.FormulaFormat.ShowImage = true;
            this.FormulaFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FormulaFormat_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.RemoveColumns);
            this.group1.Label = "Столбцы";
            this.group1.Name = "group1";
            // 
            // RemoveColumns
            // 
            this.RemoveColumns.Image = global::XAdd.Properties.Resources.deletecolumn;
            this.RemoveColumns.Label = "Удалить столбцы";
            this.RemoveColumns.Name = "RemoveColumns";
            this.RemoveColumns.ShowImage = true;
            this.RemoveColumns.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RemoveColumns_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.AppendSheets);
            this.group2.Items.Add(this.AppendSheetsCustom);
            this.group2.Items.Add(this.TableOfContents);
            this.group2.Items.Add(this.SheetsManager);
            this.group2.Items.Add(this.toggleButton1);
            this.group2.Label = "Листы";
            this.group2.Name = "group2";
            // 
            // AppendSheets
            // 
            this.AppendSheets.Image = global::XAdd.Properties.Resources.appendtable;
            this.AppendSheets.Label = "Объединить все листы в книге";
            this.AppendSheets.Name = "AppendSheets";
            this.AppendSheets.ShowImage = true;
            this.AppendSheets.SuperTip = "Будет создан лист Job, на который скопируются все листы текущей книги";
            this.AppendSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AppendSheets_Click);
            // 
            // AppendSheetsCustom
            // 
            this.AppendSheetsCustom.Image = global::XAdd.Properties.Resources.combine;
            this.AppendSheetsCustom.Label = "Объединить листы выборочно";
            this.AppendSheetsCustom.Name = "AppendSheetsCustom";
            this.AppendSheetsCustom.ShowImage = true;
            this.AppendSheetsCustom.SuperTip = "Будет создана новая книга с листом Job, на который скопируются выбранные вами лис" +
    "ты";
            this.AppendSheetsCustom.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AppendSheetsCustom_Click);
            // 
            // TableOfContents
            // 
            this.TableOfContents.Label = "Создать оглавление книги";
            this.TableOfContents.Name = "TableOfContents";
            this.TableOfContents.OfficeImageId = "TableOfContentsDialog";
            this.TableOfContents.ScreenTip = "Будет создан лист TableOfContents, на котором будет оглавление книги с ссылками н" +
    "а листы ";
            this.TableOfContents.ShowImage = true;
            this.TableOfContents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TableOfContents_Click);
            // 
            // SheetsManager
            // 
            this.SheetsManager.Label = "Диспетчер листов";
            this.SheetsManager.Name = "SheetsManager";
            this.SheetsManager.OfficeImageId = "BibliographyGallery";
            this.SheetsManager.ShowImage = true;
            this.SheetsManager.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SheetsManager_Click);
            // 
            // toggleButton1
            // 
            this.toggleButton1.Image = global::XAdd.Properties.Resources.eye_icon_png_viewed_accomms_10;
            this.toggleButton1.Label = "Показать скрытые листы";
            this.toggleButton1.Name = "toggleButton1";
            this.toggleButton1.ShowImage = true;
            this.toggleButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton1_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.InsertDate);
            this.group3.Label = "Дата";
            this.group3.Name = "group3";
            // 
            // InsertDate
            // 
            this.InsertDate.Label = "Вставить дату";
            this.InsertDate.Name = "InsertDate";
            this.InsertDate.OfficeImageId = "CalendarInsert";
            this.InsertDate.ShowImage = true;
            this.InsertDate.SuperTip = "Выделите ячейку или диапозон ячеек для вставки даты";
            this.InsertDate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InsertDate_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.Currency);
            this.group4.Label = "Числа";
            this.group4.Name = "group4";
            // 
            // Currency
            // 
            this.Currency.Image = global::XAdd.Properties.Resources.img_202966;
            this.Currency.Label = "Курсы валют";
            this.Currency.Name = "Currency";
            this.Currency.ShowImage = true;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RemoveColumns;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AppendSheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InsertDate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AppendSheetsCustom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TableOfContents;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SheetsManager;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Currency;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton FormulaFormat;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
