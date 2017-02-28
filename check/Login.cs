using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CCWin;

namespace check
{
    public partial class Login : Skin_Mac
    {
        public Login()
        {
            InitializeComponent();

        }
        
        private void skinButton1_Click(object sender, EventArgs e)
        {
            checkLogin();
        }

        DataRow dr;
        
        private void checkLogin() 
        {
            dr = check.SQL.SQL.Login(skinTextBox2.Text.ToString().Trim(),skinTextBox1.Text.ToString());
            if (dr != null)
            {
                this.DialogResult = DialogResult.OK;
                
                Main m = new Main(dr);
                this.Visible = false;
                m.ShowDialog();           
            }
            else 
            {
                skinLabel1.Text = "用户名或密码错误！";
            }

        }



        protected override bool ProcessDialogKey(Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                checkLogin();
            }
            return base.ProcessDialogKey(keyData);
        }
















        
    }
}
