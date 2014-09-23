using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Ric.Db.Info;

namespace Ric.Tasks
{
    public partial class UnderlyingAdd : Form
    {
        public UnderlyingAdd()
        {
            InitializeComponent();
        }

        public static TWUnderlyingNameInfo Prompt(string ric, string chineseDisplay, string underlyingCode)
        {
            UnderlyingAdd form = new UnderlyingAdd();
            form.tbRIC.Text = ric;
            form.tbUnderlyingCode.Text = underlyingCode;
            form.tbChineseDisplay.Text = chineseDisplay;
            form.ShowDialog();
            if (form.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                string code = form.tbCode.Text.Trim();
                string orgName = form.tbOrgName.Text.Trim();
                string engName = form.tbEngName.Text.Trim();
                string chineseChain = form.tbChineseChain.Text.Trim();
                TWUnderlyingNameInfo underlying = new TWUnderlyingNameInfo();
                underlying.UnderlyingRIC = code;
                underlying.OrganizationName = orgName;
                underlying.ChineseChain = chineseChain;
                underlying.ChineseDisplay = chineseDisplay;
                underlying.EnglishDisplay = engName;
                return underlying;
            }
            return null;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            
            List<TextBox> list = CheckTextBoxEmpty(tbCode, tbOrgName, tbEngName, tbChineseChain);
           
            if (list.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                foreach (TextBox txt in list)
                {
                    sb.Append(txt.Tag + ",");
                }
                string msg =  Convert.ToString(sb) + " Can not be empty!";
                this.ErrorMsg.Text = msg;
                this.ErrorMsg.Visible = true;               
            }

            else
            {
                this.DialogResult = DialogResult.OK;
            }
        }

        /// <summary>
        /// 检查文本框是否为空
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static List<TextBox> CheckTextBoxEmpty(params TextBox[] list)
        {
            List<TextBox> nullList = new List<TextBox>(4);
            foreach (TextBox txt in list)
            {
                if (string.IsNullOrEmpty(txt.Text.Trim()))
                {
                    nullList.Add(txt);
                }
            }
            return nullList;
        }
    }
}
