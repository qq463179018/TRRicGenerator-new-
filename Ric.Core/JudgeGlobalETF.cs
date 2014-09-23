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
    public partial class JudgeGlobalETF : Form
    {
        public JudgeGlobalETF()
        {
            InitializeComponent();
        }

        public static string Prompt(string time, string title)
        {
            JudgeGlobalETF form = new JudgeGlobalETF();           
            form.tbTime.Text = time;
            form.tbTitle.Text = title;
            form.rbNo.Checked = true;
            form.rbYes.Checked = false;
            form.ShowDialog();
            if (form.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                if (form.rbYes.Checked)
                {
                    return "Y";
                }
                return "N";
            }
            return null;
        }       
    }
}
