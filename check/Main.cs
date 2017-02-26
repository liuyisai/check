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
    public partial class Main : Skin_Mac
    {
        public Main()
        {
            InitializeComponent();
        }

        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripTextBox2_Click(object sender, EventArgs e)
        {
            
        }

        private void skinDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void skinbutton1_Click(object sender, EventArgs e)
        {
            Check ch = new Check();
            ch.Show();
        }

        private void skinButton2_Click(object sender, EventArgs e)
        {
            Count co = new Count();
            co.Show();
        }

        private void Main_Load(object sender, EventArgs e)
        {
        
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //系统时间
            toolStripTextBox4.Text ="当前时间:" +DateTime.Now.ToLocalTime().ToString();


        }

        private void refresh() 
        {
            skinDataGridView1.Rows.Clear();
                int identityCode;
                int state;
                string identityName="";
            string stateName="";
            string checkTime="";
            skinTextBox1.Text = "";
            skinComboBox3.SelectedIndex = 0;
            skinComboBox4.SelectedIndex = 0;
                string meetTime = skinDateTimePicker1.Text;
                string meetID = this.skinComboBox1.SelectedValue.ToString();
                DataTable dt2 = check.SQL.SQL.getMeeter(meetID);
                if (dt2!=null)
                {

                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        identityCode=(int)dt2.Rows[i]["identityEum"];
                        switch(identityCode)
                        {
                            case 1 :identityName="特邀代表";
                                break;

                            case 2: identityName="列席代表";
                                break;
                            case 3: identityName="正式代表";
                                break ;
                            case 0: identityName="";
                                break;

                        }
                        state=(int)dt2.Rows[i]["attendState"];
                        switch(state)
                        {
                            case 0:
                                {stateName ="否";
                                checkTime="";
                                break;
                                }
                            case 1:{stateName="是";
                            checkTime = dt2.Rows[i]["attendTime"].ToString();
                                break;}

                        }
                        skinDataGridView1.Rows.Add(dt2.Rows[i]["uId"], dt2.Rows[i]["delegationName"].ToString(), identityName,stateName,checkTime );
                        
                    }
                    
                }
            
        }



        private void skinComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (skinComboBox1.Tag == "1")
            {
                refresh(); 
                 
            }
        }

        private void skinDateTimePicker1_SelectedValueChange(object sender, string Item)
        {

            
            
            string selectTime = skinDateTimePicker1.Text;
            DataTable dt1 = check.SQL.SQL.getMeet(selectTime);
            if (dt1 != null)
            {
                this.skinComboBox1.Tag = "0";
                this.skinComboBox1 .DataSource = dt1;
                this.skinComboBox1.ValueMember = "id";
                this.skinComboBox1.DisplayMember = "mName";
                this.skinComboBox1.Tag = "1";
                refresh();
            }
            else
            {
                this.skinComboBox1.Tag = "0";
                this.skinComboBox1.DataSource = null;
                this.skinComboBox1.Tag = "1";
                this.skinDataGridView1.Rows.Clear();
                skinTextBox1.Text = "";
                skinComboBox3.SelectedIndex = 0;
                skinComboBox4.SelectedIndex = 0;
            }
        }

        private void skinTextBox1_Paint(object sender, PaintEventArgs e)
        {
            for (int i = 0; i < skinDataGridView1.Rows.Count;i++ )
            {
                int j = skinDataGridView1.Rows[i].Cells[1].Value.ToString().IndexOf(skinTextBox1.Text);
                if (j > -1)
                {
                    skinDataGridView1.Rows[i].Visible = true;
                }
                else 
                {
                    skinDataGridView1.Rows[i].Visible = false;
                }
                


            }
        }

        protected override bool ProcessDialogKey(Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                for (int i = 0; i < skinDataGridView1.Rows.Count; i++)
                {
                    int j = skinDataGridView1.Rows[i].Cells[1].Value.ToString().IndexOf(skinTextBox1.Text);
                    if (j > -1 && (skinComboBox3.SelectedItem.ToString() == skinDataGridView1.Rows[i].Cells[2].Value.ToString() || skinComboBox3.SelectedIndex == 0) && (skinComboBox4.SelectedItem.ToString() == skinDataGridView1.Rows[i].Cells[3].Value.ToString()||skinComboBox4.SelectedIndex==0))
                    {
                        skinDataGridView1.Rows[i].Visible = true;

                    }
                    else
                    {
                        skinDataGridView1.Rows[i].Visible = false;
                    }
                }
            }
            return base.ProcessDialogKey(keyData);
        }

        private void skinComboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < skinDataGridView1.Rows.Count; i++)
            {
                int a = skinDataGridView1.Rows[i].Cells[1].Value.ToString().IndexOf(skinTextBox1.Text);
                int j = skinDataGridView1.Rows[i].Cells[2].Value.ToString().IndexOf(skinComboBox3.SelectedItem.ToString());
                if (j > -1 && a > -1 && (skinComboBox4.SelectedItem.ToString() == skinDataGridView1.Rows[i].Cells[3].Value.ToString() || skinComboBox4.SelectedIndex == 0))
                {
                    skinDataGridView1.Rows[i].Visible = true;
                }
                else if (skinComboBox3.SelectedIndex == 0 && a > -1 && (skinComboBox4.SelectedItem.ToString() == skinDataGridView1.Rows[i].Cells[3].Value.ToString() || skinComboBox4.SelectedIndex == 0))
                {

                    skinDataGridView1.Rows[i].Visible = true;


                }
                else
                {
                    skinDataGridView1.Rows[i].Visible = false;
                }


                
            }
        }

        private void skinComboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < skinDataGridView1.Rows.Count; i++)
            {
                int a = skinDataGridView1.Rows[i].Cells[1].Value.ToString().IndexOf(skinTextBox1.Text);
                int j = skinDataGridView1.Rows[i].Cells[3].Value.ToString().IndexOf(skinComboBox4.SelectedItem.ToString());
                if (j > -1 && a>-1 && (skinComboBox3.SelectedItem.ToString() == skinDataGridView1.Rows[i].Cells[2].Value.ToString() || skinComboBox3.SelectedIndex == 0))
                {
                    skinDataGridView1.Rows[i].Visible = true;
                }

                else if (skinComboBox4.SelectedIndex == 0 && a>-1  && (skinComboBox3.SelectedItem.ToString() == skinDataGridView1.Rows[i].Cells[2].Value.ToString() || skinComboBox3.SelectedIndex == 0))
                {

                    skinDataGridView1.Rows[i].Visible = true;


                }
                else
                {
                    skinDataGridView1.Rows[i].Visible = false;
                }

            }
            //if (skinTextBox1.Text != "")
            //{
            //    for (int i = 0; i < skinDataGridView1.Rows.Count; i++)
            //    {
            //        int j = skinDataGridView1.Rows[i].Cells[1].Value.ToString().IndexOf(skinTextBox1.Text);
            //        if (j > -1 && skinDataGridView1.Rows[i].Cells[2].Value.ToString() == skinComboBox3.SelectedItem.ToString() && skinDataGridView1.Rows[i].Cells[3].Value.ToString() == skinComboBox4.SelectedItem.ToString())
            //        {
            //            skinDataGridView1.Rows[i].Visible = true;

            //        }
            //        else
            //        {
            //            skinDataGridView1.Rows[i].Visible = false;
            //        }
            //    }
            //}
            //else
            //{
            //    for (int i = 0; i < skinDataGridView1.Rows.Count; i++)
            //    {
                   
            //        if ( skinDataGridView1.Rows[i].Cells[2].Value.ToString() == skinComboBox3.SelectedItem.ToString() && skinDataGridView1.Rows[i].Cells[3].Value.ToString() == skinComboBox4.SelectedItem.ToString())
            //        {
            //            skinDataGridView1.Rows[i].Visible = true;

            //        }
            //        else
            //        {
            //            skinDataGridView1.Rows[i].Visible = false;
            //        }
            //    }
            //}

        }
      
          

















    }
}
