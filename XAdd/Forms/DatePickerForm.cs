using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XAdd
{
    public partial class DatePickerForm : Form
    {
        public event Action DateSelected;
        public DateTime DateSelect { get; set; }

        public DatePickerForm()
        {
            InitializeComponent();
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            Hide();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void MonthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            DateSelect = monthCalendar1.SelectionStart;
            DateSelected?.Invoke();
        }

        private void MonthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            DateSelect = monthCalendar1.SelectionStart;
            DateSelected?.Invoke();
        }
    }
}
