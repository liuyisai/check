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
    public partial class Check :Skin_Mac
    {
        string MeetId;
        public delegate void DelegateInsertMain();

        public event DelegateInsertMain ChangeTimer;  
        public Check(string meetId)
        {
            InitializeComponent();
            MeetId = meetId;

        }
        private static Check instance;

        public static Check CreateForm(string meetId)
        {
            if (instance == null || instance.IsDisposed)
            {
                instance = new Check(meetId);

            }
            return instance;
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
             checking();                            
        }

        private void checking()
        {
            try
            {
                if (PingIpOrDomainName("115.24.161.31"))
                {
                    string QRcode = skinTextBox1.Text.ToString().Trim();
                    string userPosition = "";
                    string userChecktime = "";
                    string identityName = "";
                    int identityCode = 0;
                    DataTable dt;
                    int i = 0;
                    userChecktime = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                    dt = check.SQL.SQL.getMeeterInfo(QRcode, MeetId);
                    if (dt != null)
                    {
                        identityCode = (int)dt.Rows[0]["identityEum"];
                        switch (identityCode)
                        {
                            case 1: identityName = "特邀代表";
                                break;
                            case 2: identityName = "列席代表";
                                break;
                            case 3: identityName = "正式代表";
                                break;
                            case 0: identityName = "";
                                break;
                        }
                        if ((int)dt.Rows[0]["attendState"] == 0)
                        {
                            i = check.SQL.SQL.setMeeterInfo(QRcode, userChecktime);
                            if (i == 1)
                            {

                                textBox1.Text = dt.Rows[0]["uName"].ToString();
                                textBox2.Text = dt.Rows[0]["delegationName"].ToString();
                                userPosition = QRcode.Substring(5, 2) + "排" + QRcode.Substring(7, 2) + "列";
                                textBox3.Text = userPosition;
                                textBox4.Text = userChecktime;
                                textBox5.Text = identityName;
                                skinTextBox1.Text = "";
                                skinLabel5.Text = "签到成功！";

                            }
                          
                        }
                        else
                        {
                            textBox1.Text = dt.Rows[0]["uName"].ToString();
                            textBox2.Text = dt.Rows[0]["delegationName"].ToString();
                            userPosition = QRcode.Substring(5, 2) + "排" + QRcode.Substring(7, 2) + "列";
                            textBox3.Text = userPosition;
                            textBox5.Text = identityName;
                            textBox4.Text = dt.Rows[0]["attendTime"].ToString();
                            skinLabel5.Text = "此人已签到！";
                            skinTextBox1.Text = "";
                        }
                    }
                    else
                    {
                        skinLabel5.Text = "查无此人！";
                        skinTextBox1.Text = "";
                        textBox1.Text = "";
                        textBox2.Text = "";
                        textBox3.Text = "";
                        textBox4.Text = "";
                        textBox5.Text = "";

                    }
                }
                else 
                {
                    MessageBox.Show("请检查网络连接！");
                }
            }
            catch (Exception)
            {
                
                MessageBox.Show("网络异常！");
                skinTextBox1.Text = "";
            }
           
           
            


            
        }
        protected override bool ProcessDialogKey(Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                checking();
            }
            return base.ProcessDialogKey(keyData);
        }

        private void Check_Load(object sender, EventArgs e)
        {

        }

        private void Check_FormClosing(object sender, FormClosingEventArgs e)
        {
            ChangeTimer();
            //Main m = new Main();
            //shutdownRefresh shut = new shutdownRefresh(m.doRefresh);
            //shut();
        }
        int flag = 0;

        private void skinSplitContainer1_SplitterMoved(object sender, SplitterEventArgs e)
        {
            if (flag==0)
            {
                flag = 1;
                this.Height = Height - 150;           
                return;
            }
            else 
            {
                flag = 0;
                this.Height = Height + 150;               
                return;
                
            }

        }

   
    }
}
