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
    public partial class InputReferenceFields : Form
    {
        public InputReferenceFields()
        {
            InitializeComponent();
        }

        public static List<string> Prompt(string edcoid)
        {
            InputReferenceFields form = new InputReferenceFields();
            form.TextBoxEdcoid.Text = edcoid;
            form.ShowDialog();
            if (form.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                List<string> result = new List<string>();
                result.Add(form.textBox1.Text.Trim());
                result.Add(form.textBox2.Text.Trim());
                return result;
            }
            return null;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            string msg = "Please delete symbols like '.'','";
            string issuerOrgId = textBox1.Text;
            string issuerLegalName = textBox2.Text;
            if (string.IsNullOrEmpty(issuerOrgId))
            {
                label5.Visible = true;
            }
            if (string.IsNullOrEmpty(issuerLegalName))
            {
                label6.Visible = true;
            }
            if (!(string.IsNullOrEmpty(issuerOrgId) || string.IsNullOrEmpty(issuerLegalName)))
            {
                if (issuerOrgId.Contains(".") || issuerOrgId.Contains(","))
                {
                    label5.Text = msg;
                    label5.Visible = true;
                }
                if (issuerLegalName.Contains(".") || issuerLegalName.Contains(","))
                {
                    label6.Text = msg;
                    label6.Visible = true;
                }
                if (!(issuerOrgId.Contains(".") || issuerOrgId.Contains(",") || issuerLegalName.Contains(".") || issuerLegalName.Contains(",")))
                {
                    this.DialogResult = DialogResult.OK;
                }
            }
        }
    }
}
