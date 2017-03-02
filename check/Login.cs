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
using System.Net.NetworkInformation;

namespace check
{
    public partial class Login : Skin_Mac
    {
        public Login()
        {
            InitializeComponent();
    

        }

        private bool PingIpOrDomainName(string strIpOrDName)
        {
            try
            {
                Ping objPingSender = new Ping();
                PingOptions objpinOptions = new PingOptions();
                objpinOptions.DontFragment = true;
                string data = "";
                byte[] buffer = Encoding.UTF8.GetBytes(data);
                int intTimeout = 120;
                PingReply objPinReply = objPingSender.Send(strIpOrDName, intTimeout, buffer, objpinOptions);
                string strInfo = objPinReply.Status.ToString();
                if (strInfo == "Success")
                {
                    return true;

                }
                else
                {
                    return false;
                }


            }
            catch (Exception)
            {

                return false;
            }

        }
        
        private void skinButton1_Click(object sender, EventArgs e)
        {
            checkLogin();
        }

        DataRow dr;
        
        private void checkLogin() 
        {
            try
            {
                if (PingIpOrDomainName("115.24.161.31"))
                {
                    dr = check.SQL.SQL.Login(skinTextBox2.Text.ToString().Trim(), skinTextBox1.Text.ToString());
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
                else 
                {
                    MessageBox.Show("请检查网络连接！");
                }
                

            }
            catch (Exception)
            {

                MessageBox.Show("请检查网络连接！");
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
