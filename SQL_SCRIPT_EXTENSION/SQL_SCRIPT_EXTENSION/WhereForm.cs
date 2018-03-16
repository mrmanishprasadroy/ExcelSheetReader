/* Author : roym
 * Date: 20/03/2016
 * Where statement Pop up  
 * 
 * */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SQL_SCRIPT_EXTENSION
{
    public partial class WhereForm : Form
    {
        public string WhereStaement = string.Empty;
        private static int AddHitCOunt = 0;
        /// <summary>
        /// Initlize
        /// </summary>
        public WhereForm()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Over ride Intilization
        /// </summary>
        /// <param name="coulmnInfo"></param>
        public WhereForm(List<string> coulmnInfo)
        {
            InitializeComponent();

            Bind(coulmnInfo);
        }
        /// <summary>
        /// Bind method 
        /// </summary>
        /// <param name="coulmnInfo"></param>
        private void Bind(List<string> coulmnInfo)
        {
            coulmnlookup.Properties.DataSource = coulmnInfo.ToList();
        }
        /// <summary>
        /// Add where condition 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            
            if (string.IsNullOrEmpty(textEdit1.Text) || string.IsNullOrEmpty(coulmnlookup.Text))
            {
                MessageBox.Show(string.Format("one of the value is empty, value: {0}, coulmn name : {1}",textEdit1.Text,coulmnlookup.Text.ToString()));
            }
            else
            {
                StringBuilder Swhere = new StringBuilder();
                AddHitCOunt++;
                if (AddHitCOunt == 1)
                {
                    
                    Swhere.Append("Where ");
                    Swhere.Append(coulmnlookup.Text);
                    Swhere.Append("= '" + textEdit1.Text);
                    Swhere.Append("'");
                    WhereStaement = Swhere.ToString();
                    textEdit1.Text = "";
                    coulmnlookup.Text = "";
                }
                else
                {
                    Swhere.Append(" AND ");
                    Swhere.Append(coulmnlookup.Text);
                    Swhere.Append("= '" + textEdit1.Text);
                    Swhere.Append("'");
                    WhereStaement += Swhere.ToString();
                    textEdit1.Text = "";
                    coulmnlookup.Text = "";
                }
                memoEdit1.Text = WhereStaement.ToString();
               // memoEdit1.Text = memoEdit1.Text + ";";

                
            }
        }
        /// <summary>
        /// OK
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(WhereStaement))
            {
                MessageBox.Show("No Where Statment was build");
                DialogResult = System.Windows.Forms.DialogResult.Cancel;
                Close();
            }
            else
            {
                WhereStaement = memoEdit1.Text.ToString();
                DialogResult = System.Windows.Forms.DialogResult.OK;
                Close();           
            }
        }
        /// <summary>
        /// Memo edit 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void memoEdit1_EditValueChanged(object sender, EventArgs e)
        {
            //WhereStaement += memoEdit1.Text;
        }
    }
}
