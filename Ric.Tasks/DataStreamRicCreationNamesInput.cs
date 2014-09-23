using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Ric.Tasks
{
    public partial class DataStreamRicCreationNamesInput : Form
    {
        private nameInputType type = nameInputType.Name;
        public DataStreamRicCreationNamesInput()
        {
            InitializeComponent();
        }
        public DataStreamRicCreationNamesInput(nameInputType inputType)
        {
            InitializeComponent();
            type = inputType;
            if (inputType == nameInputType.Name)
            {
                name1_label.Text = "NAME1:";
                name2_label.Text = "NAME2:";
            }
            else
            {
                name1_label.Text = "FNAME1:";
                name2_label.Text = "FNAME2:";
            }
        }
        public static string[] Prompt(string companyName, string abbName, nameInputType inputType = nameInputType.Name)
        {
            DataStreamRicCreationNamesInput form = new DataStreamRicCreationNamesInput(inputType);
            form.tbCompanyName.Text = companyName;
            form.tbAbbName.Text = abbName;
            form.lbWarning.Visible = false;
            form.ShowDialog();
            if (form.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                string[] names = new string[2];
                names[0] = form.tbName1.Text.Trim();
                names[1] = form.tbName2.Text.TrimEnd();
                return names;
            }
            return null;
        }

        public static string[] Prompt(string companyName, string abbName)
        {
            DataStreamRicCreationNamesInput form = new DataStreamRicCreationNamesInput();
            form.tbCompanyName.Text = companyName;
            form.tbAbbName.Text = abbName;
            form.lbWarning.Visible = false;
            form.ShowDialog();
            if (form.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                string[] names = new string[2];
                names[0] = form.tbName1.Text.Trim();
                names[1] = form.tbName2.Text.TrimEnd();                
                return names;
            }
            return null;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            lbWarning.Visible = false;
            string name1 = tbName1.Text.Trim();
            string name2 = tbName2.Text.TrimEnd();

            if (name1.Length > 24)
            {
                lbWarning.Text = "NAME1 must <= 24 characters!";
                lbWarning.Visible = true;
            }

            else if (name2.Length > 24)
            {
                lbWarning.Text = "NAME2 must <= 24 characters!";
                lbWarning.Visible = true;
            }

            else
            {             
                this.DialogResult = DialogResult.OK;
            }
        }       
    }
}
