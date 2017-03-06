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
using System.Threading;

namespace check
{
    public partial class Check :Skin_Mac
    {
        string MeetId;
        string userId;
        DataTable allResource;
       // public delegate void DelegateInsertMain();

       // public event DelegateInsertMain ChangeTimer;  
        public Check(string meetId,string userID,DataTable dta)
        {
            InitializeComponent();
            MeetId = meetId;
            userId = userID;
            allResource = dta;
            
        }
        private static Check instance;

        public static Check CreateForm(string meetId,string userID,DataTable dta)
        {
            if (instance == null || instance.IsDisposed)
            {
                instance = new Check(meetId,userID,dta);

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
                
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                if (PingIpOrDomainName("115.24.161.31"))
                {
                    Function();
                }
                else 
                {
                    Thread.Sleep(3000);
                    if (PingIpOrDomainName("115.24.161.31"))
                    {
                        Function();
                    }
                    else 
                    {
                        MessageBox.Show("请检查网络连接！");
                    }
                   
                }
            }
            catch (Exception)
            {
                
                MessageBox.Show("网络异常！");
                skinTextBox1.Text = "";
            }
           
           
            


            
        }
        int personNum = 0;

        private void Function() 
        {
            string QRcode = skinTextBox1.Text.ToString().Trim();
            string userPosition = "";
            string userChecktime = "";
            string identityName = "";
            int identityCode = 0;
            DataTable dt;
            int i = 0;
            string year = DateTime.Now.ToString("yyyy-MM-dd");
            string time = DateTime.Now.ToString("HH:mm:ss");
            userChecktime = year + " " + time;
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
                    i = check.SQL.SQL.setMeeterInfo(QRcode, userChecktime, userId);
                    if (i == 1)
                    {

                        textBox1.Text = dt.Rows[0]["uName"].ToString();
                        textBox2.Text = dt.Rows[0]["delegationName"].ToString();
                        userPosition = QRcode.Substring(5, 2) + "排" + QRcode.Substring(7, 2) + "列";
                        textBox3.Text = userPosition;
                        textBox4.Text = time;
                        textBox5.Text = identityName;
                        skinTextBox1.Text = "";
                        skinLabel5.Text = "签到成功！";


                        #region 上部数据表
                        skinDataGridView1.Rows.Insert(0, textBox1.Text.ToString(), textBox2.Text.ToString(), textBox5.Text.ToString(), userChecktime);
                        skinDataGridView1.Rows[1].Selected = false;
                        skinDataGridView1.Rows[0].Selected = true;
                        personNum++;
                        skinLabel7.Text = "签到口-流量统计：" + personNum.ToString();
                        #endregion
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
            DataView view =allResource.DefaultView;
            view.Sort="Column7 DESC";
            DataTable dtall=view.ToTable();
            //skinDataGridView1.Sort(skinDataGridView1.Columns["签到时间"],ListSortDirection.Descending);
            for (int i = 0; i < dtall.Rows.Count; i++)
            {
                if (dtall.Rows[i]["Column3"].ToString()==userId.ToString())
                {
                    skinDataGridView1.Rows.Add(dtall.Rows[i]["Column1"], dtall.Rows[i]["Column5"], dtall.Rows[i]["Column2"], dtall.Rows[i]["Column7"]);
                    personNum++;
                    
                }
                
            }
            //skinDataGridView1.Sort(col);
            skinLabel7.Text = "签到口-流量统计："+personNum.ToString();
        }

        private void Check_FormClosing(object sender, FormClosingEventArgs e)
        {
            //ChangeTimer();
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
                this.Width = Width - 590;           
                return;
            }
            else 
            {
                flag = 0;
                this.Width = Width + 590;               
                return;
                
            }

        }

   
    }
}
