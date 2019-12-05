using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace XAdd
{
    public partial class ThisAddIn
    {
        #region Переменные
        Office.CommandBar cb = null;
        Office.CommandBarButton buttonContext = null;
        DatePickerForm form_DatePicker = new DatePickerForm();
        AppendSheetsForm form_AppendSheetsCustom = new AppendSheetsForm();
        SheetsManagerForm form_SheetsManager = new SheetsManagerForm();
        SheetRenameForm form_SheetRename = new SheetRenameForm();
        CurrencyForm form_Currency = new CurrencyForm();
        List<string> sheetsName = new List<string>();
        long lastRow;
        long lastCol;
        bool answer = false;
        Excel.Range area;
        string shName;
        Random rnd = new Random();

        #endregion
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            #region Обработчики_ОбъединениеЛистов
            form_AppendSheetsCustom.checkBox2.CheckedChanged += form_AppendSheetsCustomCheckBox2_CheckedChanged;
            form_AppendSheetsCustom.SelectedNodesToFinal.Click += AppendSheetsCustom_SelectedNodesToList;
            form_AppendSheetsCustom.treeView1.DoubleClick += AppendSheetsCustom_SelectedNodesToList;
            form_AppendSheetsCustom.RemoveNodesFromFinal.Click += AppendSheetsCustom_RemoveNodesFromList;
            form_AppendSheetsCustom.treeView2.DoubleClick += AppendSheetsCustom_RemoveNodesFromList;
            form_AppendSheetsCustom.AppendSheetsOK.Click += AppendSheetsCustom_Append;

            #endregion

            #region Обработчики_ДиспетчерЛистов

            form_SheetsManager.SheetsManagerClickNode += SheetsManagerClickNode;
            form_SheetsManager.SheetsManagerDoubleClickNode += SheetsManagerDoubleClickNode;
            form_SheetsManager.SheetsManagerOpenClicked += Form_SheetsManager_SheetsManagerOpen;
            form_SheetsManager.SheetsManagerRenameClicked += Form_SheetsManager_SheetsManagerRename;
            form_SheetsManager.SheetsManagerRemoveClicked += Form_SheetsManager_SheetsManagerRemove;
            form_SheetsManager.SheetsManagerNewBookClicked += Form_SheetsManager_SheetsManagerNewBook;
            form_SheetsManager.SheetsManagerNewSheetClicked += Form_SheetsManager_SheetsManagerNewSheet;
            form_SheetsManager.SheetsManagerCreateCopyClicked += Form_SheetsManager_SheetsManagerCreateCopy;

            #endregion


            form_DatePicker.DateSelected += DatePicker_dateSelected;// обработчик выбор даты

            cb = Application.CommandBars["Cell"]; 
            buttonContext = cb.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true) as Office.CommandBarButton; // Кнопка "Протянуть формулу"
            buttonContext.Caption = "Протянуть формулу (XAdd)";
            buttonContext.Tag = "FormulaFil";
            buttonContext.Style = Office.MsoButtonStyle.msoButtonCaption;
            buttonContext.Click += ButtonContext_Click;
            buttonContext.Visible = true;
            

            buttonContext = cb.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true) as Office.CommandBarButton; // Кнопка "Заменить формулы на значения"
            buttonContext.Caption = "Заменить формулы на значения (XAdd)";
            buttonContext.Tag = "ReplaceFormulasWithValues";
            buttonContext.Style = Office.MsoButtonStyle.msoButtonCaption;
            buttonContext.Click += ReplaceFormulasWithValues;
            buttonContext.Visible = true;

        }



        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {

            var ribbon = new Ribbon1();
            ribbon.ButtonRemoveColumnsClicked += Ribbon_ButtonRemoveColumns;
            ribbon.ButtonAppendSheetsClicked += Ribbon_ButtonAppendSheets;
            ribbon.ButtonInsertDateClicked += Ribbon_ButtonInsertDate;
            ribbon.ButtonAppendSheetsCustom += Ribbon_ButtonAppendSheetsCustom;
            ribbon.ButtonTableOfContentsClicked += Ribbon_ButtonTableOfContents;
            ribbon.ButtonSheetsManagerClicked += Ribbon_ButtonSheetsManager;
            ribbon.ButtonShowHiddenSheetsClicked += Ribbon_ButtonShowHiddenSheets;
            ribbon.ButtonHideHiddenSheetsClicked += Ribbon_ButtonHideHiddenSheets;
            ribbon.ButtonCurrencyClicked += Ribbon_ButtonCurrency;
            ribbon.ButtonFormulaFormatEnableClicked += Ribbon_ButtonFormulaFormatEnable;
            ribbon.ButtonFormulaFormatDisableClicked += Ribbon_ButtonFormulaFormatDisable;
            ribbon.ButtonShowSheetsShortcutsClicked += Ribbon_ButtonShowSheetsShortcuts;
            ribbon.ButtonHideSheetsShortcutsClicked += Ribbon_ButtonHideSheetsShortcuts;
            ribbon.ButtonAutoFillClicked += Ribbon_ButtonAutoFill;
            ribbon.ButtonCalculatorClicked += Ribbon_ButtonCalculator;
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { ribbon });
        }










        #region Удаление столбцов
        private void Ribbon_ButtonRemoveColumns() //удаляет столбцы на активном листе. кнопка нажата
        {

            Excel.Worksheet activeSheet = this.Application.ActiveSheet;

            Excel.Range FindCell;
            Excel.Range Row;
            int RowNumber;
            bool ExactMatch = true;

            int ColumnsRemovedCount = 0;
            if (int.TryParse(Microsoft.VisualBasic.Interaction.InputBox("Введите номер строки с названиями столбцов (числовое значение)", "XAdd"), out int rowInput) && rowInput > 0)
            {
                RowNumber = rowInput;
            }
            else
            {
                Microsoft.VisualBasic.Interaction.MsgBox("Введите правильное значение!", Microsoft.VisualBasic.MsgBoxStyle.Exclamation, "XAdd");
                return;
            }


            string SearchName = Microsoft.VisualBasic.Interaction.InputBox("Введите название столбца (без кавычек)", "XAdd", "");

            var responseExactMatch = Microsoft.VisualBasic.Interaction.MsgBox("Применить точное совпадение?", Buttons: Microsoft.VisualBasic.MsgBoxStyle.YesNoCancel, "XAdd");
            switch (responseExactMatch)
            {
                case Microsoft.VisualBasic.MsgBoxResult.Cancel:
                    return;
                case Microsoft.VisualBasic.MsgBoxResult.Yes:
                    ExactMatch = true;
                    break;
                case Microsoft.VisualBasic.MsgBoxResult.No:
                    ExactMatch = false;
                    break;
                default:
                    break;
            }


            if (!string.IsNullOrWhiteSpace(SearchName))
            {
                if (ExactMatch)
                {
                    foreach (Excel.Range column in activeSheet.Columns)
                    {
                        Row = activeSheet.Rows[RowNumber];


                        FindCell = Row.Cells.Find(SearchName, LookIn: Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlWhole);


                        if (FindCell != null)
                        {
                            FindCell.EntireColumn.Delete();
                            ColumnsRemovedCount = ++ColumnsRemovedCount;
                        }
                        else
                        {

                        }

                    }
                }
                else
                {
                    foreach (Excel.Range column in activeSheet.Columns)
                    {
                        Row = activeSheet.Rows[RowNumber];


                        FindCell = Row.Cells.Find(SearchName);


                        if (FindCell != null)
                        {
                            FindCell.EntireColumn.Delete();
                            ColumnsRemovedCount = ++ColumnsRemovedCount;
                        }
                        else
                        {

                        }

                    }
                }

                if (ColumnsRemovedCount == 0)
                {
                    Microsoft.VisualBasic.Interaction.MsgBox("Ничего не найдено!", Microsoft.VisualBasic.MsgBoxStyle.Exclamation, "XAdd");
                    return;
                }

            }
            else
            {
                Microsoft.VisualBasic.Interaction.MsgBox("Пустой запрос!", Microsoft.VisualBasic.MsgBoxStyle.Exclamation, "XAdd");
                return;

            }
            Microsoft.VisualBasic.Interaction.MsgBox(string.Format("Успешно удалено {0} ст.", ColumnsRemovedCount), Microsoft.VisualBasic.MsgBoxStyle.Information, "XAdd");

        }
        #endregion

        #region Объединение листов
        private void Ribbon_ButtonAppendSheets() // объединяет все листы в активной книге. кнопка нажата
        {
            Application.DisplayAlerts = false;

            DialogResult dr = MessageBox.Show("Нужно ли копировать формулы? (в случае отрицательного ответа будут скопированы только значения)", "XAdd", MessageBoxButtons.YesNoCancel);
            switch (dr)
            {
                case DialogResult.None:
                    return;
                case DialogResult.OK:
                    answer = true;
                    break;
                case DialogResult.Cancel:
                    return;
                case DialogResult.Yes:
                    answer = true;
                    break;
                case DialogResult.No:
                    answer = false;
                    break;
                default:
                    break;
            }

            try
            {
                Application.Sheets["Job"].Delete();
            }
            catch (Exception)
            {

            }


            Application.Sheets.Add(Before: Application.Sheets[1], Count: 1);
            Application.ActiveSheet.Name = "Job";
            Excel.Worksheet jobSheet = Application.Sheets["Job"];

            foreach (Excel.Worksheet ws in Application.Sheets)
            {
                if (ws.Index != 1)
                {
                    try
                    {
                        lastCol = ws.Cells.Find("*", System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns,
                        Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                    }
                    catch (Exception)
                    {

                        continue;
                    }

                    lastRow = ws.Cells.Find("*", System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows,
                    Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                    shName = "*********************** " + ws.Name + " ******************************";
                    ws.Range[ws.Cells[1, 1], ws.Cells[lastRow, lastCol]].Copy();
                    lastRow = jobSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    area = Application.Cells[lastRow, 1].MergeArea;
                    if (area.Cells.Count > 1)
                    {
                        lastRow += area.Cells.Count;
                    }
                    else
                    {
                        lastRow += 1;
                    }
                    jobSheet.Cells[lastRow, 1].EntireRow.Interior.ColorIndex = 6;
                    jobSheet.Cells[lastRow, 1].Value = shName;
                    if (answer)
                    {
                        jobSheet.Paste(jobSheet.Cells[lastRow + 1, 1]);
                    }
                    else
                    {
                        jobSheet.Cells[lastRow + 1, 1].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, missing, missing);
                        jobSheet.Cells[lastRow + 1, 1].PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, missing, missing);
                    }


                }

            }
            jobSheet.Cells[1, 1].EntireRow.Delete();
            Application.DisplayAlerts = true;
        }

        #endregion

        #region Выбор даты
        private void Ribbon_ButtonInsertDate() // показывает пользователю форму с календарем. нажата кнопка на риббоне ( см. форму DatePickerForm )
        {

            form_DatePicker.StartPosition = FormStartPosition.CenterScreen;
            form_DatePicker.Show();
        }

        //private void ribbon_ButtonInsertDate(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault) // показывает пользователю форму с календарем. нажата кнопка в контекстном меню
        //{
        //    Point mousePoint = new Point(Cursor.Position.X, Cursor.Position.Y);
        //    form_DatePicker.Location = mousePoint;
        //    form_DatePicker.StartPosition = FormStartPosition.Manual;
        //    form_DatePicker.Show();
        //}



        private void DatePicker_dateSelected() // вставляет дату, выбранную в календаре( см. форму DatePickerForm )
        {

            DateTime datePicked = form_DatePicker.DateSelect;


            Application.ActiveWindow.RangeSelection.Cells.NumberFormat = "m/d/yyyy";
            Excel.Range selectedRange = Application.ActiveWindow.RangeSelection.Cells;

            if (selectedRange?.Columns.Count == 1 && selectedRange?.Count > 1)
            {
                selectedRange[1].Value = datePicked;
                selectedRange[1].AutoFill(selectedRange, Excel.XlAutoFillType.xlFillDefault);
            }
            else
            {
                foreach (Excel.Range cell in Application.ActiveWindow.RangeSelection.Cells)
                {
                    cell.Value = datePicked;
                    datePicked = datePicked.AddDays(1);
                }
            }

            selectedRange.Columns.AutoFit();

            form_DatePicker.Hide();

        }


        #endregion

        #region Кастомное объединение листов

        private void form_AppendSheetsCustomCheckBox2_CheckedChanged(object sender, EventArgs e) // обработчик события: нажат чекбокс отображать скрытые листы
        {
            form_AppendSheetsCustomFillNode();
        }

        private void form_AppendSheetsCustomFillNode() // заполняет список книг/листов
        {
            form_AppendSheetsCustom.treeView1.Nodes.Clear();
            if (form_AppendSheetsCustom.checkBox2.Checked)
            {
                foreach (Excel.Workbook wb in Application.Workbooks)
                {
                    TreeNode tempWorkbookNode = form_AppendSheetsCustom.treeView1.Nodes.Add(wb.Name, wb.Name);
                    tempWorkbookNode.BackColor = Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256));
                    TreeNode[] tnd = form_AppendSheetsCustom.treeView1.Nodes.Find(wb.Name, false);
                    form_AppendSheetsCustom.treeView1.SelectedNode = tnd[0];
                    foreach (Excel.Worksheet ws in wb.Sheets)
                    {
                        TreeNode tempSheetNode = form_AppendSheetsCustom.treeView1.SelectedNode.Nodes.Add(wb.Name, ws.Name);
                        tempSheetNode.BackColor = tempWorkbookNode.BackColor;
                    }
                }
            }
            else
            {
                foreach (Excel.Workbook wb in Application.Workbooks)
                {
                    TreeNode tempWorkbookNode = form_AppendSheetsCustom.treeView1.Nodes.Add(wb.Name, wb.Name);
                    tempWorkbookNode.BackColor = Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256));
                    TreeNode[] tnd = form_AppendSheetsCustom.treeView1.Nodes.Find(wb.Name, false);
                    form_AppendSheetsCustom.treeView1.SelectedNode = tnd[0];
                    foreach (Excel.Worksheet ws in wb.Sheets)
                    {
                        if (ws.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                        {
                            TreeNode tempSheetNode = form_AppendSheetsCustom.treeView1.SelectedNode.Nodes.Add(wb.Name, ws.Name);
                            tempSheetNode.BackColor = tempWorkbookNode.BackColor;
                        }
                    }
                }
            }
        }

        private void Ribbon_ButtonAppendSheetsCustom() // наполнение Treeview1 из списка открытых книг
        {
            form_AppendSheetsCustom.treeView1.Nodes.Clear();
            form_AppendSheetsCustom.treeView2.Nodes.Clear();
            form_AppendSheetsCustomFillNode();
            form_AppendSheetsCustom.Show();

        }

        private void AppendSheetsCustom_SelectedNodesToList(object sender, System.EventArgs e) // клонирование выбранных книг/листов из Treeview1  в Treeview2
        {
            try
            {
                TreeNode clonedNode = (TreeNode)form_AppendSheetsCustom.treeView1.SelectedNode.Clone();
                form_AppendSheetsCustom.treeView2.Nodes.Add(clonedNode);
            }
            catch (Exception)
            {

                return;
            }

        }

        private void AppendSheetsCustom_RemoveNodesFromList(object sender, System.EventArgs e) // удаление выбранных книг/листов из Treeview2
        {
            try
            {
                form_AppendSheetsCustom.treeView2.SelectedNode.Remove();
            }
            catch (Exception)
            {

                return;
            }

        }

        private void AppendSheetsCustom_Append(object sender, System.EventArgs e) // объединение листов (кнопка нажата)
        {
            if (form_AppendSheetsCustom.checkBox1.Checked)
            { // объединение с учетом заголовков

                DialogResult dr = MessageBox.Show("Нужно ли копировать формулы? (в случае отрицательного ответа будут скопированы только значения)", "XAdd", MessageBoxButtons.YesNoCancel);
                switch (dr)
                {
                    case DialogResult.None:
                        return;
                    case DialogResult.OK:
                        answer = true;
                        break;
                    case DialogResult.Cancel:
                        return;
                    case DialogResult.Yes:
                        answer = true;
                        break;
                    case DialogResult.No:
                        answer = false;
                        break;
                    default:
                        break;
                }

                Application.Workbooks.Add();
                string jobWbString = Application.ActiveWorkbook.Name;
                Application.ActiveSheet.Name = "Job";
                form_AppendSheetsCustom.treeView2.SelectedNode = form_AppendSheetsCustom.treeView2.Nodes[0];
                if (form_AppendSheetsCustom.treeView2.SelectedNode.Nodes.Count > 0)
                {
                    form_AppendSheetsCustom.treeView2.SelectedNode = form_AppendSheetsCustom.treeView2.SelectedNode.FirstNode;
                }
                Excel.Workbook actWb = Application.Workbooks.Item[form_AppendSheetsCustom.treeView2.SelectedNode.Name];
                Excel.Worksheet actSheet = actWb.Sheets[form_AppendSheetsCustom.treeView2.SelectedNode.Text];
                lastCol = actSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                actSheet.Range[actSheet.Cells[1, 1], actSheet.Cells[1, lastCol]].Copy();
                Excel.Workbook jobWb = Application.Workbooks.Item[jobWbString];
                Excel.Worksheet jobSheet = jobWb.Sheets["Job"];
                jobSheet.Paste(jobSheet.Cells[1, 1]);
                foreach (TreeNode node in form_AppendSheetsCustom.treeView2.Nodes)
                {

                    if (node.Nodes.Count > 0)
                    {
                        foreach (TreeNode childNode in node.Nodes)
                        {
                            actWb = Application.Workbooks.Item[childNode.Name];
                            actSheet = actWb.Sheets[childNode.Text];

                            try
                            {
                                lastCol = actSheet.Cells.Find("*", System.Reflection.Missing.Value,
                                System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns,
                                Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                            }
                            catch (Exception)
                            {

                                continue;
                            }


                            lastRow = actSheet.Cells.Find("*", System.Reflection.Missing.Value,
                            System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows,
                            Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                            actSheet.Range[actSheet.Cells[2, 1], actSheet.Cells[lastRow, lastCol]].Copy();
                            lastRow = jobSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                            area = Application.Cells[lastRow, 1].MergeArea;
                            if (area.Cells.Count > 1)
                            {
                                lastRow += area.Cells.Count;
                            }
                            else
                            {
                                lastRow += 1;
                            }
                            if (answer)
                            {
                                jobSheet.Paste(jobSheet.Cells[lastRow, 1]);
                            }
                            else
                            {
                                jobSheet.Cells[lastRow, 1].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, missing, missing);
                                jobSheet.Cells[lastRow, 1].PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, missing, missing);
                            }
                        }

                    }
                    else
                    {
                        actWb = Application.Workbooks.Item[node.Name];
                        actSheet = actWb.Sheets[node.Text];

                        try
                        {
                            lastCol = actSheet.Cells.Find("*", System.Reflection.Missing.Value,
                            System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns,
                            Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                        }
                        catch (Exception)
                        {

                            continue;
                        }


                        lastRow = actSheet.Cells.Find("*", System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows,
                        Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                        actSheet.Range[actSheet.Cells[2, 1], actSheet.Cells[lastRow, lastCol]].Copy();
                        lastRow = jobSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        area = Application.Cells[lastRow, 1].MergeArea;
                        if (area.Cells.Count > 1)
                        {
                            lastRow += area.Cells.Count;
                        }
                        else
                        {
                            lastRow += 1;
                        }
                        if (answer)
                        {
                            jobSheet.Paste(jobSheet.Cells[lastRow, 1]);
                        }
                        else
                        {
                            jobSheet.Cells[lastRow, 1].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, missing, missing);
                            jobSheet.Cells[lastRow, 1].PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, missing, missing);
                        }
                    }
                }
            }
            else //Полное объединение (checkbox не отмечен)
            {


                DialogResult dr = MessageBox.Show("Нужно ли копировать формулы? (в случае отрицательного ответа будут скопированы только значения)", "XAdd", MessageBoxButtons.YesNoCancel);
                switch (dr)
                {
                    case DialogResult.None:
                        return;
                    case DialogResult.OK:
                        answer = true;
                        break;
                    case DialogResult.Cancel:
                        return;
                    case DialogResult.Yes:
                        answer = true;
                        break;
                    case DialogResult.No:
                        answer = false;
                        break;
                    default:
                        break;
                }

                Application.Workbooks.Add();
                string JobWb = Application.ActiveWorkbook.Name;
                Application.ActiveSheet.Name = "Job";

                foreach (TreeNode node in form_AppendSheetsCustom.treeView2.Nodes)
                {

                    if (node.Nodes.Count > 0)
                    {
                        foreach (TreeNode childNode in node.Nodes)
                        {
                            Excel.Workbook actWb = Application.Workbooks.Item[childNode.Name];
                            Excel.Worksheet actSheet = actWb.Sheets[childNode.Text];

                            try
                            {
                                lastCol = actSheet.Cells.Find("*", System.Reflection.Missing.Value,
                                System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns,
                                Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                            }
                            catch (Exception)
                            {

                                continue;
                            }


                            lastRow = actSheet.Cells.Find("*", System.Reflection.Missing.Value,
                            System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows,
                            Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                            shName = "*********************** " + actWb.Name + "\\" + actSheet.Name + " ******************************";
                            actSheet.Range[actSheet.Cells[1, 1], actSheet.Cells[lastRow, lastCol]].Copy();
                            Excel.Workbook jobWb = Application.Workbooks.Item[JobWb];
                            Excel.Worksheet jobSheet = jobWb.Sheets["Job"];
                            lastRow = jobSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                            area = Application.Cells[lastRow, 1].MergeArea;
                            if (area.Cells.Count > 1)
                            {
                                lastRow += area.Cells.Count;
                            }
                            else
                            {
                                lastRow += 1;
                            }
                            jobSheet.Cells[lastRow, 1].EntireRow.Interior.ColorIndex = 6;
                            jobSheet.Cells[lastRow, 1].Value = shName;
                            if (answer)
                            {
                                jobSheet.Paste(jobSheet.Cells[lastRow + 1, 1]);
                            }
                            else
                            {
                                jobSheet.Cells[lastRow + 1, 1].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, missing, missing);
                                jobSheet.Cells[lastRow + 1, 1].PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, missing, missing);
                            }
                        }

                    }
                    else
                    {
                        Excel.Workbook actWb = Application.Workbooks.Item[node.Name];
                        Excel.Worksheet actSheet = actWb.Sheets[node.Text];

                        try
                        {
                            lastCol = actSheet.Cells.Find("*", System.Reflection.Missing.Value,
                            System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns,
                            Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                        }
                        catch (Exception)
                        {

                            continue;
                        }

                        lastRow = actSheet.Cells.Find("*", System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows,
                        Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                        shName = "*********************** " + actWb.Name + "\\" + actSheet.Name + " ******************************";
                        actSheet.Range[actSheet.Cells[1, 1], actSheet.Cells[lastRow, lastCol]].Copy();
                        Excel.Workbook jobWb = Application.Workbooks.Item[JobWb];
                        Excel.Worksheet jobSheet = jobWb.Sheets["Job"];
                        lastRow = jobSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        area = Application.Cells[lastRow, 1].MergeArea;
                        if (area.Cells.Count > 1)
                        {
                            lastRow += area.Cells.Count;
                        }
                        else
                        {
                            lastRow += 1;
                        }
                        jobSheet.Cells[lastRow, 1].EntireRow.Interior.ColorIndex = 6;
                        jobSheet.Cells[lastRow, 1].Value = shName;
                        if (answer)
                        {
                            try
                            {
                                jobSheet.Paste(jobSheet.Cells[lastRow + 1, 1]);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message,"XAdd");
                            }
                            
                        }
                        else
                        {
                            try
                            {
                                jobSheet.Cells[lastRow + 1, 1].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, missing, missing);
                                jobSheet.Cells[lastRow + 1, 1].PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, missing, missing);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message,"XAdd");
                            }
                            
                        }
                    }
                }
                Excel.Worksheet finishSheet = Application.Sheets["Job"];
                finishSheet.Cells[1, 1].EntireRow.Delete();

            }
            Clipboard.Clear();
        }



        #endregion

        #region Оглавление книги

        private void Ribbon_ButtonTableOfContents()
        {
            //int LastCol;
            Application.
            Application.DisplayAlerts = false;

            try
            {
                Application.Sheets["Оглавление"].Delete();
            }
            catch (Exception)
            {

            }

            Application.Sheets.Add(Before: Application.Sheets[1], Count: 1);
            Application.ActiveSheet.Name = "Оглавление";
            Excel.Worksheet jobSheet = Application.Sheets["Оглавление"];
            Excel.Range startCell = jobSheet.Cells[2, 1];
            //Excel.Range pictureCell = jobSheet.Cells[2, 2];
            jobSheet.Cells[1, 1].Value = "Оглавление книги";
            jobSheet.Cells[1, 1].Font.Size = 20;


            foreach (Excel.Worksheet ws in Application.Worksheets)
            {
                if (ws.Index != 1)
                {
                    jobSheet.Hyperlinks.Add(startCell, "", ws.Name + "!A1", missing, ws.Name);
                    startCell = startCell.Offset[1, 0];
                    // превью листов
                    //try
                    //{
                    //    LastCol = ws.Cells.Find("*", System.Reflection.Missing.Value,
                    //        System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns,
                    //        Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                    //}
                    //catch (Exception)
                    //{

                    //    continue;
                    //}
                    //ws.Range[ws.Cells[1, 1], ws.Cells[100, LastCol]].CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap); //вывод превью листа
                    //jobSheet.Paste(pictureCell, Clipboard.GetImage());
                    //pictureCell = pictureCell.Offset[1, 0];

                }
            }

            Application.DisplayAlerts = true;


        }

        #endregion

        #region Диспетчер листов

        private void Ribbon_ButtonSheetsManager()
        {
            form_SheetsManager.Hide();
            form_SheetsManager.treeView1.Nodes.Clear(); //наполнение Treeview списком книг и листов

            foreach (Excel.Workbook wb in Application.Workbooks)
            {
                form_SheetsManager.treeView1.Nodes.Add(wb.Name, wb.Name);
                TreeNode[] tnd = form_SheetsManager.treeView1.Nodes.Find(wb.Name, false);
                form_SheetsManager.treeView1.SelectedNode = tnd[0];
                foreach (Excel.Worksheet ws in wb.Sheets)
                {
                    form_SheetsManager.treeView1.SelectedNode.Nodes.Add(wb.Name, ws.Name);
                }
            }
            form_SheetsManager.treeView1.ExpandAll();
            form_SheetsManager.Show();
        }

        private void SheetsManagerDoubleClickNode()
        { //двойной клик по листу из Treeview1
            if (form_SheetsManager.treeView1.SelectedNode.Parent != null)
            {
                Excel.Workbook actWb = Application.Workbooks.Item[form_SheetsManager.treeView1.SelectedNode.Name];
                Excel.Worksheet actSheet = actWb.Sheets.Item[form_SheetsManager.treeView1.SelectedNode.Text];
                actSheet.Activate();

            }
        }

        private void SheetsManagerClickNode() //клик по листу из Treeview1
        {
            int lastCol = 1;

            if (form_SheetsManager.treeView1.SelectedNode.Parent != null)
            {
                Excel.Workbook actWb = Application.Workbooks.Item[form_SheetsManager.treeView1.SelectedNode.Name];
                Excel.Worksheet actSheet = actWb.Sheets.Item[form_SheetsManager.treeView1.SelectedNode.Text];


                try
                {
                    lastCol = actSheet.Cells.Find("*", System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns,
                        Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                }
                catch (Exception)
                {

                }
                actSheet.Range[actSheet.Cells[1, 1], actSheet.Cells[100, lastCol]].CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap); //вывод превью листа
                form_SheetsManager.pictureBox1.Image = Clipboard.GetImage();
                Clipboard.Clear();
            }
        }

        private void Form_SheetsManager_SheetsManagerRemove() // кнопка удаление листа
        {
            if (form_SheetsManager.treeView1.SelectedNode.Parent != null)
            {

                try
                {
                    Excel.Workbook actWb;
                    Excel.Worksheet actSheet;

                    foreach (TreeNode node in form_SheetsManager.treeView1.Nodes)
                    {
                        foreach (TreeNode childNode in node.Nodes)
                        {
                            if (childNode.Checked)
                            {
                                actWb = Application.Workbooks.Item[childNode.Name];
                                actSheet = actWb.Sheets.Item[childNode.Text];
                                actSheet.Delete();
                            }
                            else if (childNode == form_SheetsManager.treeView1.SelectedNode)
                            {
                                actWb = Application.Workbooks.Item[childNode.Name];
                                actSheet = actWb.Sheets.Item[childNode.Text];
                                actSheet.Delete();
                            }

                        }

                    }

                    Form_SheetsManager_Refresh();
                    form_SheetsManager.treeView1.SelectedNode = form_SheetsManager.treeView1.SelectedNode.LastNode;
                    form_SheetsManager.treeView1.Focus();
                }
                catch (Exception ex)
                {
                    Form_SheetsManager_Refresh();
                    MessageBox.Show(ex.Message, "XAdd", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }

            }
        }

        private void Form_SheetsManager_SheetsManagerRename() // кнопка переименовать лист
        {

            int countChecked = 0;

            foreach (TreeNode node in form_SheetsManager.treeView1.Nodes)
            {
                foreach (TreeNode childNode in node.Nodes)
                {
                    if (childNode.Checked)
                    {
                        countChecked++;
                    }
                    if (countChecked > 1)
                    {
                        break;
                    }
                }
            }


            if (form_SheetsManager.treeView1.SelectedNode.Parent != null && countChecked < 2)
            {

                Excel.Workbook actWb = Application.Workbooks.Item[form_SheetsManager.treeView1.SelectedNode.Name];
                Excel.Worksheet actSheet = actWb.Sheets.Item[form_SheetsManager.treeView1.SelectedNode.Text];
                form_SheetRename.SetSheetName(actSheet.Name);
                DialogResult dr = form_SheetRename.ShowDialog();

                if (dr == DialogResult.OK)
                {
                    try
                    {
                        actSheet.Name = form_SheetRename.textBox1.Text;
                        Form_SheetsManager_Refresh();

                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message, "XAdd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }



            }
            else
            {
                MessageBox.Show("Выберите один лист!", "XAdd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void Form_SheetsManager_SheetsManagerOpen() //кнопка открыть книгу
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files (*.xls,*.xls,*.xlsm,*.xla,*.xlsb,*.xlam)|*.xl*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Application.Workbooks.Open(ofd.FileName);
                Form_SheetsManager_Refresh();
                form_SheetsManager.Activate();
            }


        }

        private void Form_SheetsManager_SheetsManagerCreateCopy() // кнопка сделать копию листа
        {
            if (form_SheetsManager.treeView1.SelectedNode.Nodes.Count == 0)
            {
                TreeNode selectedNode = form_SheetsManager.treeView1.SelectedNode;
                Excel.Workbook actWb = Application.Workbooks.Item[form_SheetsManager.treeView1.SelectedNode.Name];
                Excel.Worksheet actSheet = actWb.Sheets.Item[form_SheetsManager.treeView1.SelectedNode.Text];
                try
                {
                    actSheet.Copy(missing, actSheet);
                    TreeNode addNode = form_SheetsManager.treeView1.SelectedNode.Parent.Nodes.Insert(selectedNode.Index + 1, actWb.Name, Application.ActiveSheet.Name);
                    form_SheetsManager.treeView1.SelectedNode = addNode;
                    form_SheetsManager.treeView1.Focus();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "XAdd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        private void Form_SheetsManager_SheetsManagerNewSheet() //кнопка добавить лист
        {


            if (form_SheetsManager.treeView1.SelectedNode.Parent != null)
            {
                TreeNode selectedNode = form_SheetsManager.treeView1.SelectedNode;
                Excel.Workbook actWb = Application.Workbooks.Item[form_SheetsManager.treeView1.SelectedNode.Name];
                Excel.Worksheet actSheet = actWb.Sheets.Item[form_SheetsManager.treeView1.SelectedNode.Text];

                try
                {
                    Excel.Worksheet newSheet = actWb.Worksheets.Add(missing, actSheet, 1, Excel.XlSheetType.xlWorksheet);
                    TreeNode addNode = form_SheetsManager.treeView1.SelectedNode.Parent.Nodes.Insert(selectedNode.Index + 1, actWb.Name, newSheet.Name);
                    form_SheetsManager.treeView1.SelectedNode = addNode;
                    form_SheetsManager.treeView1.Focus();
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message, "XAdd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                Excel.Workbook actWb = Application.Workbooks.Item[form_SheetsManager.treeView1.SelectedNode.Name];
                TreeNode selectedNode = form_SheetsManager.treeView1.SelectedNode;
                try
                {
                    Excel.Worksheet newSheet = actWb.Worksheets.Add(missing, missing, 1, Excel.XlSheetType.xlWorksheet);
                    TreeNode addNode = form_SheetsManager.treeView1.SelectedNode.Nodes.Insert(selectedNode.Index + 1, actWb.Name, newSheet.Name);
                    form_SheetsManager.treeView1.SelectedNode = addNode;
                    form_SheetsManager.treeView1.Focus();
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message, "XAdd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }
        private void Form_SheetsManager_SheetsManagerNewBook() //кнопка добавить книгу
        {
            Application.Workbooks.Add(missing);
            Form_SheetsManager_Refresh();
            form_SheetsManager.Activate();
        }

        private void Form_SheetsManager_Refresh()// обновление Treeview книг и листов
        {

            form_SheetsManager.treeView1.Nodes.Clear();

            foreach (Excel.Workbook wb in Application.Workbooks)
            {
                form_SheetsManager.treeView1.Nodes.Add(wb.Name, wb.Name);
                TreeNode[] tnd = form_SheetsManager.treeView1.Nodes.Find(wb.Name, false);
                form_SheetsManager.treeView1.SelectedNode = tnd[0];
                foreach (Excel.Worksheet ws in wb.Sheets)
                {
                    form_SheetsManager.treeView1.SelectedNode.Nodes.Add(wb.Name, ws.Name);
                }
            }
            form_SheetsManager.treeView1.ExpandAll();
        }

        #endregion

        #region Показать/скрыть скрытые листы
        private void Ribbon_ButtonShowHiddenSheets()
        {


            foreach (Excel.Worksheet ws in Application.ActiveWorkbook.Sheets)
            {
                if (ws.Visible == Excel.XlSheetVisibility.xlSheetHidden || ws.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden)
                {
                    sheetsName.Add(ws.Name);
                    ws.Tab.Color = Color.PaleVioletRed;
                    ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                }

            }
        }

        private void Ribbon_ButtonHideHiddenSheets()
        {
            foreach (string sheetName in sheetsName)
            {
                Application.Worksheets[sheetName].Visible = Excel.XlSheetVisibility.xlSheetHidden;
            }
            sheetsName.Clear();

        }
        #endregion

        #region Курсы валют
        private void Ribbon_ButtonCurrency() // курсы валют кнопка нажата
        {
            form_Currency.Hide();

            form_Currency.Show();
        }

        #endregion

        #region Формат формул
        private void Ribbon_ButtonFormulaFormatEnable()
        {
            Application.ReferenceStyle = Excel.XlReferenceStyle.xlR1C1;
        }

        private void Ribbon_ButtonFormulaFormatDisable()
        {
            Application.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
        }

        #endregion

        #region Показывать панель листов
        private void Ribbon_ButtonHideSheetsShortcuts()
        {
            Application.ActiveWindow.DisplayWorkbookTabs = false;
        }

        private void Ribbon_ButtonShowSheetsShortcuts()
        {
            Application.ActiveWindow.DisplayWorkbookTabs = true;
        }

        #endregion

        #region Тестовая функция умное автозаполнение
        private void Ribbon_ButtonAutoFill()
        {
            Excel.Range activeCell = Application.ActiveCell;

            try
            {
                lastRow = Application.ActiveSheet.Cells.Find("*", System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows,
                    Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            }
            catch (Exception)
            {

                return;
            }

            Application.ActiveSheet.Cells[lastRow, activeCell.Column].Value = 1;



            Excel.Range wRange = Application.Range[Application.ActiveCell, Application.ActiveCell.EntireColumn.Find("*", System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns,
                        Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Cells]
                .SpecialCells(Excel.XlCellType.xlCellTypeVisible, missing);

            wRange.Cells.Value = activeCell.FormulaR1C1;

            //foreach (Excel.Range cell in wRange)
            //{
            //    cell.FormulaR1C1 = activeCell.FormulaR1C1;
            //}

        }

        private void ButtonContext_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Ribbon_ButtonAutoFill();
        }
        #endregion

        #region Калькулятор

        private void Ribbon_ButtonCalculator()
        {
            Application.ActivateMicrosoftApp(Index: 0);
        }

        #endregion

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        #region Замена формул на значения

        private void ReplaceFormulasWithValues(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Excel.Range selRange = Application.ActiveWindow.RangeSelection.Cells;
            selRange.Copy();
            selRange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            Clipboard.Clear();
        }

        #endregion

    }
}