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
    public partial class IssuerAdd : Form
    {
        public IssuerAdd()
        {
            InitializeComponent();
        }

        public static TWIssueInfo Prompt(string ric, string warrantName, string orgName)
        {
            IssuerAdd form = new IssuerAdd();
            form.tbRIC.Text = ric;
            form.tbWarrantName.Text = warrantName;
            form.tbOrgName.Text = orgName;
            form.ShowDialog();
            if (form.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                string chineseShortName = form.tbChineseShortName.Text.Trim();
                string englishBriefName = form.tbEnglishBriefName.Text.Trim();
                string englishShortName = form.tbEnglishShortName.Text.Trim();
                string englishName = form.tbEnglishName.Text.Trim();
                string englishFullName = form.tbEnglishFullName.Text.Trim();
                string issuerCode = form.tbIssuerCode.Text.Trim();
                string warrantIssuer = form.tbWarrantIssuer.Text.Trim();
                string chineseChain = form.tbWarrantName.Text.Trim();
                TWIssueInfo issuer = new TWIssueInfo();
                issuer.ChineseShortName = chineseShortName;
                issuer.EnglishBriefName = englishBriefName;
                issuer.EnglishShortName = englishShortName;
                issuer.EnglishName = englishName;
                issuer.EnglishFullName = englishFullName;
                issuer.IssueCode = issuerCode;
                issuer.WarrantIssuer = warrantIssuer;
                return issuer;
            }
            return null;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            
            List<TextBox> list = CheckTextBoxEmpty(tbChineseShortName, tbEnglishBriefName, tbEnglishShortName, tbEnglishName, tbEnglishFullName, tbIssuerCode, tbWarrantIssuer);
           
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
