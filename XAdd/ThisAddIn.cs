using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Drawing;

namespace XAdd
{
    public partial class ThisAddIn
    {
        DatePickerForm form_DatePicker = new DatePickerForm();
        AppendSheetsForm form_AppendSheetsCustom = new AppendSheetsForm();
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            form_AppendSheetsCustom.SelectedNodesToFinalList += AppendSheetsCustom_SelectedNodesToList;
            form_AppendSheetsCustom.RemoveNodesFromFinalList += AppendSheetsCustom_RemoveNodesFromList;
            form_AppendSheetsCustom.AppendSheetsClicked += AppendSheetsCustom_Append;
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
            long LastRow;
            long LastCol;
            string shName;

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

                    LastRow = ws.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    LastCol = ws.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    shName = "*********************** " + ws.Name + " ******************************";
                    ws.Range[ws.Cells[1, 1], ws.Cells[LastRow, LastCol]].Copy();
                    LastRow = jobSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                    jobSheet.Cells[LastRow, 1].EntireRow.Interior.ColorIndex = 6;
                    jobSheet.Cells[LastRow, 1].Value = shName;
                    jobSheet.Paste(jobSheet.Cells[LastRow + 1, 1]);

                }

            }
            Microsoft.VisualBasic.Interaction.MsgBox("Листы успешно объединены!", Buttons: Microsoft.VisualBasic.MsgBoxStyle.Information, "XAdd");
            Application.DisplayAlerts = true;
        }

        #endregion

        #region Выбор даты
        private void Ribbon_ButtonInsertDate() // показывает пользователю форму с календарем. нажата кнопка на риббоне
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

        private void DatePicker_dateSelected() // вставляет дату, выбранную в календаре
        {

            DateTime datePicked = form_DatePicker.DateSelect;

            foreach (Excel.Range cell in Application.ActiveWindow.RangeSelection.Cells)
            {
                cell.Value = datePicked;
                cell.NumberFormat = "m/d/yyyy";
                datePicked = datePicked.AddDays(1);
            }

            form_DatePicker.Hide();
        }


        #endregion

        #region Кастомное объединение листов

        private void Ribbon_ButtonAppendSheetsCustom() // наполнение Treeview1 из списка открытых книг
        {
            form_AppendSheetsCustom.treeView1.Nodes.Clear();
            form_AppendSheetsCustom.treeView2.Nodes.Clear();

            foreach (Excel.Workbook wb in Application.Workbooks)
            {
                form_AppendSheetsCustom.treeView1.Nodes.Add(wb.Name,wb.Name);
                TreeNode[] tnd = form_AppendSheetsCustom.treeView1.Nodes.Find(wb.Name, false);
                form_AppendSheetsCustom.treeView1.SelectedNode = tnd[0];
                foreach (Excel.Worksheet ws in wb.Sheets)
                {

                    form_AppendSheetsCustom.treeView1.SelectedNode.Nodes.Add(wb.Name, ws.Name);

                }
            }

            form_AppendSheetsCustom.Show();

        }

        private void AppendSheetsCustom_SelectedNodesToList() // клонирование выбранных книг/листов из Treeview1  в Treeview2
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

        private void AppendSheetsCustom_RemoveNodesFromList() // удаление выбранных книг/листов из Treeview2
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

        private void AppendSheetsCustom_Append() // объединение листов (кнопка нажата)
        {
            if (form_AppendSheetsCustom.checkBox1.Checked)
            { // объединение с учетом заголовков
                long LastRow;
                long LastCol;
                Application.Workbooks.Add();
                string jobWbString = Application.ActiveWorkbook.Name;
                Application.ActiveSheet.Name = "Job";
                form_AppendSheetsCustom.treeView2.SelectedNode = form_AppendSheetsCustom.treeView2.Nodes[0];
                if (form_AppendSheetsCustom.treeView2.SelectedNode.Nodes.Count > 0)
                {
                    form_AppendSheetsCustom.treeView2.SelectedNode = form_AppendSheetsCustom.treeView2.SelectedNode.FirstNode;
                }
                Excel.Workbook actWb = Application.Workbooks.Item[form_AppendSheetsCustom.treeView2.SelectedNode.Name];
                Excel.Worksheet actSheet= actWb.Sheets[form_AppendSheetsCustom.treeView2.SelectedNode.Text];
                LastCol = actSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                actSheet.Range[actSheet.Cells[1, 1],actSheet.Cells[1, LastCol]].Copy();
                Excel.Workbook jobWb = Application.Workbooks.Item[jobWbString];
                Excel.Worksheet jobSheet = jobWb.Sheets["Job"];
                jobSheet.Paste(jobSheet.Cells[1,1]);
                foreach (TreeNode node in form_AppendSheetsCustom.treeView2.Nodes)
                {

                    if (node.Nodes.Count > 0)
                    {
                        foreach (TreeNode childNode in node.Nodes)
                        {
                            actWb = Application.Workbooks.Item[childNode.Name];
                            actSheet = actWb.Sheets[childNode.Text];
                            LastRow = actSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                            LastCol = actSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                            actSheet.Range[actSheet.Cells[2, 1], actSheet.Cells[LastRow, LastCol]].Copy();
                            LastRow = jobSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                            jobSheet.Paste(jobSheet.Cells[LastRow, 1]);
                        }

                    }
                    else
                    {
                        actWb = Application.Workbooks.Item[node.Name];
                        actSheet = actWb.Sheets[node.Text];
                        LastRow = actSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        LastCol = actSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                        actSheet.Range[actSheet.Cells[2, 1], actSheet.Cells[LastRow, LastCol]].Copy();
                        LastRow = jobSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                        jobSheet.Paste(jobSheet.Cells[LastRow, 1]);
                    }
                }
            }
            else //Полное объединение (checkbox не отмечен)
            {
                long LastRow;
                long LastCol;
                string shName;
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
                            LastRow = actSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                            LastCol = actSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                            shName = "*********************** " + actWb.Name + "\\" + actSheet.Name + " ******************************";
                            actSheet.Range[actSheet.Cells[1, 1], actSheet.Cells[LastRow, LastCol]].Copy();
                            Excel.Workbook jobWb = Application.Workbooks.Item[JobWb];
                            Excel.Worksheet jobSheet = jobWb.Sheets["Job"];
                            LastRow = jobSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                            jobSheet.Cells[LastRow, 1].EntireRow.Interior.ColorIndex = 6;
                            jobSheet.Cells[LastRow, 1].Value = shName;
                            jobSheet.Paste(jobSheet.Cells[LastRow + 1, 1]);
                        }

                    }
                    else
                    {
                        Excel.Workbook actWb = Application.Workbooks.Item[node.Name];
                        Excel.Worksheet actSheet = actWb.Sheets[node.Text];
                        LastRow = actSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        LastCol = actSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                        shName = "*********************** " + actWb.Name + "\\" + actSheet.Name + " ******************************";
                        actSheet.Range[actSheet.Cells[1, 1], actSheet.Cells[LastRow, LastCol]].Copy();
                        Excel.Workbook jobWb = Application.Workbooks.Item[JobWb];
                        Excel.Worksheet jobSheet = jobWb.Sheets["Job"];
                        LastRow = jobSheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                        jobSheet.Cells[LastRow, 1].EntireRow.Interior.ColorIndex = 6;
                        jobSheet.Cells[LastRow, 1].Value = shName;
                        jobSheet.Paste(jobSheet.Cells[LastRow + 1, 1]);
                    }
                }
            } 
            
        }



        #endregion

        #region Оглавление книги

        private void Ribbon_ButtonTableOfContents()
        {
            Application.DisplayAlerts = false;

            try
            {
                Application.Sheets["TableOfContents"].Delete();
            }
            catch (Exception)
            {

            }

            Application.Sheets.Add(Before: Application.Sheets[1], Count: 1);
            Application.ActiveSheet.Name = "TableOfContents";
            Excel.Worksheet jobSheet = Application.Sheets["TableOfContents"];
            Excel.Range startCell = jobSheet.Cells[2,1];
            jobSheet.Cells[1, 1].Value = "Table of Contents";

            foreach (Excel.Worksheet ws in Application.Worksheets)
            {
                if (ws.Index!=1)
                {
                    jobSheet.Hyperlinks.Add(startCell, "", ws.Name+"!A1", missing, ws.Name);
                    startCell = startCell.Offset[1,0];
                }
            }

            Application.DisplayAlerts = true;


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


      
    }
}