﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace XAdd
{
    public partial class CurrencyForm : Form
    {
        public string dateSelected { get; set; }
        public CurrencyForm()
        {
            InitializeComponent();
        }

        private void CurrencyForm_Load(object sender, EventArgs e)
        {
            monthCalendar1.SelectionStart = DateTime.Now;

        }

        private void CurrencyForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            Hide();
        }


        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            label1.Text = "Курс на "+ monthCalendar1.SelectionStart.ToString("dd'/'MM'/'yyyy");
            dataGridView1.Rows.Clear();
            dataGridView1.Rows.Add();
            if (monthCalendar1.SelectionStart > DateTime.Now)
            {
                monthCalendar1.SelectionStart = DateTime.Now;
                monthCalendar1.SelectionEnd = DateTime.Now;
            }
            dateSelected = monthCalendar1.SelectionStart.ToString("dd'/'MM'/'yyyy");
            XmlDocument xDoc = new XmlDocument();
            try
            {
                xDoc.Load(string.Format(@"http://www.cbr.ru/scripts/XML_daily.asp?date_req={0}", dateSelected));
            }
            catch (Exception)
            {

                MessageBox.Show("Нет доступа к интернету или источник недоступен!","XAdd",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }
            
            XmlElement xRoot = xDoc.DocumentElement;

            foreach (XmlNode xnode in xRoot)
            {
                if (xnode.Attributes.Count > 0)
                {
                    XmlNode attr = xnode.Attributes.GetNamedItem("ID");

                    if (attr != null)
                    {
                        if (attr.Value == "R01235")
                        {
                            foreach (XmlNode childNode in xnode.ChildNodes)
                            {
                                if (childNode.Name == "Value")
                                {
                                    dataGridView1.Rows[0].Cells["USD"].Value = childNode.InnerText;
                                }
                            }
                        }
                        if (attr.Value == "R01239")
                        {
                            foreach (XmlNode childNode in xnode.ChildNodes)
                            {
                                if (childNode.Name == "Value")
                                {
                                    dataGridView1.Rows[0].Cells["EUR"].Value = childNode.InnerText;
                                }
                            }
                        }
                        if (attr.Value == "R01035")
                        {
                            foreach (XmlNode childNode in xnode.ChildNodes)
                            {
                                if (childNode.Name == "Value")
                                {
                                    dataGridView1.Rows[0].Cells["GBP"].Value = childNode.InnerText;
                                }
                            }
                        }


                    }
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.cbr.ru");
        }
    }
}
