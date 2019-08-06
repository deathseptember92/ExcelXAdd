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
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

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
    }
}
