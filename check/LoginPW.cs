using CCWin;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace check
{
    public partial class LoginPW : Skin_Mac
    {
        public LoginPW()
        {
            InitializeComponent();
            
        }

        private void LoginPW_Load(object sender, EventArgs e)
        {

        }

        private void skinButton2_Click(object sender, EventArgs e)
        {
            if (skinTextBox3.Text .Trim ()==skinTextBox4.Text .Trim())
            {
                int i = SQL.SQL.updatePassword(skinTextBox2.Text .Trim (),skinTextBox1.Text .Trim(),skinTextBox3.Text.Trim ());
                if (i > 0)
                {
                    MessageBox.Show("修改成功!");
                    this.DialogResult = DialogResult.OK;
                }
                else { MessageBox.Show("账号或密码输入有误!"); }
            }
            else
            {
                MessageBox.Show("输入的两次密码不同！");
            }
        }
    }
}
