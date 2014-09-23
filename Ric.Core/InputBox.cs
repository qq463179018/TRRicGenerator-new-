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
    public partial class InputBox : Form
    {
        public InputBox()
        {
            InitializeComponent();
        }

        public static string Prompt(string title, string msg)
        {
            return Prompt(title, msg, false);
        }

        public static string Prompt(string title, string msg, bool showCombo)
        {
            InputBox form = new InputBox();
            form.Text = string.IsNullOrEmpty(title) ? "Input Box" : title;
            form.textBox2.Text = string.IsNullOrEmpty(msg) ? "Please input:" : msg;
            form.comboBox1.Visible = showCombo;
            form.comboBox1.SelectedIndex = 0;

            form.ShowDialog();

            if (form.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                if (showCombo)
                {
                    return string.Format("{0},{1}", form.textBox1.Text.Trim(), form.comboBox1.SelectedItem.ToString());
                }

                return form.textBox1.Text.Trim();
            }

            return null;
        }
    }
}
