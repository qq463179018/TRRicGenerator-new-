using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace Ric.Core
{
    public partial class InputISINType : Form
    {
        public InputISINType()
        {
            InitializeComponent();
        }

        public static string Prompt(string koreaName)
        {
            InputISINType form = new InputISINType();
            form.textBox2.Text = koreaName;
            form.ShowDialog();
            if (form.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                string result = form.textBox1.Text.Trim();
                string type = string.Empty;
                if (form.radioButton1.Checked)
                {
                    type = "ORD";
                }
                else if (form.radioButton2.Checked)
                {
                    type = "PRF";
                }
                else
                {
                    type = "KDR";
                }

                return result + "," + type;
            }
            return null;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            string isin = textBox1.Text;
            label4.Visible = false;
            if (string.IsNullOrEmpty(isin))
            {
                label4.Visible = true;
                return;
            }
            Regex regex = new Regex(@"[^A-Z0-9]+");
            Match match = regex.Match(isin);
            if (match.Success)
            {
                label4.Text = "Bad format of ISIN!";
                label4.Visible = true;
                return;
            }
            
            this.DialogResult = DialogResult.OK;

        }
    }
}
