using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace XAdd
{
    public partial class Ribbon1
    {
        public event Action ButtonRemoveColumnsClicked;
        public event Action ButtonAppendSheetsClicked;
        public event Action ButtonInsertDateClicked;
        public event Action ButtonAppendSheetsCustom;
        public event Action ButtonTableOfContentsClicked;
        public event Action ButtonSheetsManagerClicked;
        public event Action ButtonShowHiddenSheetsClicked;
        public event Action ButtonHideHiddenSheetsClicked;
        public event Action ButtonCurrencyClicked;
        public event Action ButtonFormulaFormatEnableClicked;
        public event Action ButtonFormulaFormatDisableClicked;
        public event Action ButtonShowSheetsShortcutsClicked;
        public event Action ButtonHideSheetsShortcutsClicked;
        public event Action ButtonAutoFillClicked;
        public event Action ButtonCalculatorClicked;
        public event Action ButtonSortSheetsClicked;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            ShowSheetsShortcuts.Checked=true;
        }

        private void RemoveColumns_Click(object sender, RibbonControlEventArgs e)
        {
            ButtonRemoveColumnsClicked?.Invoke();
        }

        private void AppendSheets_Click(object sender, RibbonControlEventArgs e)
        {
            ButtonAppendSheetsClicked?.Invoke();
        }


        private void InsertDate_Click(object sender, RibbonControlEventArgs e)
        {
            ButtonInsertDateClicked?.Invoke();
        }

        private void AppendSheetsCustom_Click(object sender, RibbonControlEventArgs e)
        {
            ButtonAppendSheetsCustom?.Invoke();
        }

        private void TableOfContents_Click(object sender, RibbonControlEventArgs e)
        {
            ButtonTableOfContentsClicked?.Invoke();

        }

        private void SheetsManager_Click(object sender, RibbonControlEventArgs e)
        {
            ButtonSheetsManagerClicked?.Invoke();
        }


        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            if (toggleButton1.Checked)
            {
                toggleButton1.Image = XAdd.Properties.Resources.show_hide_password_10_512;
                toggleButton1.Label = "Скрыть листы";
                ButtonShowHiddenSheetsClicked?.Invoke();
            }
            else
            {
                toggleButton1.Image = XAdd.Properties.Resources.eye_icon_png_viewed_accomms_10;
                toggleButton1.Label = "Показать скрытые листы";
                ButtonHideHiddenSheetsClicked?.Invoke();
            }
        }

        private void FormulaFormat_Click(object sender, RibbonControlEventArgs e)
        {
            if (FormulaFormat.Checked)
            {
                ButtonFormulaFormatEnableClicked?.Invoke();
            }
            else
            {
                ButtonFormulaFormatDisableClicked?.Invoke();
            }
            
        }

        private void Currency_Click(object sender, RibbonControlEventArgs e)
        {
            ButtonCurrencyClicked?.Invoke();
        }

        private void ShowSheetsShortcuts_Click(object sender, RibbonControlEventArgs e)
        {
            if (ShowSheetsShortcuts.Checked)
            {
                ButtonShowSheetsShortcutsClicked?.Invoke();
            }
            else
            {
                ButtonHideSheetsShortcutsClicked?.Invoke();
            }
        }

        private void AutoFill_Click(object sender, RibbonControlEventArgs e)
        {
            ButtonAutoFillClicked?.Invoke();
        }

        private void Calculator_Click(object sender, RibbonControlEventArgs e)
        {
            ButtonCalculatorClicked?.Invoke();
        }

        private void SortSheets_Click(object sender, RibbonControlEventArgs e)
        {
            ButtonSortSheetsClicked?.Invoke();
        }
    }
}
