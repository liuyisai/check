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
    public partial class Check :Skin_Mac
    {
        string MeetId;
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









        private void skinButton1_Click(object sender, EventArgs e)
         {


             checking();


            
            
            

        }

        private void checking()
        {
            string QRcode = skinTextBox1.Text.ToString().Trim();
            string userPosition = "";
            string userChecktime = "";
            DataTable dt;
            userChecktime = DateTime.Now.ToLocalTime().ToString();

            int i = check.SQL.SQL.setMeeterInfo(QRcode, userChecktime);


            if (i == 1)
            {
                dt = check.SQL.SQL.getMeeterInfo(QRcode, MeetId);
                textBox1.Text = dt.Rows[0]["uName"].ToString();
                textBox2.Text = dt.Rows[0]["delegationName"].ToString();
                userPosition = QRcode.Substring(5, 2) + "排" + QRcode.Substring(7, 2) + "列";
                textBox3.Text = userPosition;
                textBox4.Text = userChecktime;
                skinTextBox1.Text = "";
                skinLabel5.Text = "签到成功！";

            }
            else
            {
                skinLabel5.Text = "查无此人！";
                skinTextBox1.Text = "";
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";

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
           
        }

   
    }
}
