using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Ric.Core
{
    public partial class InputISIN : Form
    {
        public InputISIN()
        {
            InitializeComponent();
        }

        public static string Prompt(string underlyingName, string title)
        {
            InputISIN form = new InputISIN();
            form.label1.Text = title + ":";
            form.textBox2.Text = underlyingName;
            form.ShowDialog();
            if (form.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                string result = form.textBox1.Text.Trim();               
                return result;
            }
            return null;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {           
            string issuerOrgId = textBox1.Text;            
            label4.Visible = false;           
            if (string.IsNullOrEmpty(issuerOrgId))
            {
                label4.Visible = true;
            }           
            else
            {               
                this.DialogResult = DialogResult.OK;
            }
        }
    }
}
