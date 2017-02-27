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
            

            if (skinComboBox1.SelectedValue!=null )
            {
                string Meet = skinComboBox1.SelectedValue.ToString();
                Check ch = Check.CreateForm(Meet);
                ch.Show();

                timer2.Enabled = true;
                if (ch.DialogResult==DialogResult.OK)
                {
                    
                }
                




            }
            else 
            {
                MessageBox.Show("无法签到，请选择会议！", "提示信息", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            }
           
        }
        private void skinButton2_Click(object sender, EventArgs e)
        {
            //Count co = new Count(MainDt);
            Count co;
            MainDt = GetDgvToTable(skinDataGridView1);
            if (MainDt.Rows .Count  != 0)
            {
                co = Count.CreateForm(MainDt);
                co.Show();

            }
            else 
            {
                MessageBox.Show("当前会议无数据，请重新选择会议！","提示信息",MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            }
           
        }
        public DataTable GetDgvToTable(DataGridView dgv)
        {
            DataTable dt = new DataTable();

            // 列强制转换
            for (int count = 0; count < dgv.Columns.Count; count++)
            {
                DataColumn dc = new DataColumn(dgv.Columns[count].Name.ToString());
                dt.Columns.Add(dc);
            }

            // 循环行
            for (int count = 0; count < dgv.Rows.Count; count++)
            {
                DataRow dr = dt.NewRow();
                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
                {
                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
        private void Main_Load(object sender, EventArgs e)
        {
        
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //系统时间
            toolStripTextBox4.Text ="当前时间:" +DateTime.Now.ToLocalTime().ToString();


        }
        DataTable MainDt=null ;
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
                System .Data .DataTable dt2 = check.SQL.SQL.getMeeter(meetID);
                if (dt2 != null)
                {

                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        identityCode = (int)dt2.Rows[i]["identityEum"];
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
                        state = (int)dt2.Rows[i]["attendState"];
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
                        skinDataGridView1.Rows.Add(dt2.Rows[i]["uName"], dt2.Rows[i]["delegationName"].ToString(), identityName, stateName, checkTime);
                        
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
           

        }

        private void skinButton3_Click(object sender, EventArgs e)
        {
            string fileName = "";
            string saveFileName = "";
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xls";
            saveDialog.Filter = "Excel文件|*.xls";
            saveDialog.FileName = fileName;
            saveDialog.ShowDialog();
            saveFileName = saveDialog.FileName;
            if (saveFileName.IndexOf(":") < 0) return; //被点了取消
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("无法创建Excel对象，您的电脑可能未安装Excel");
                return;
            }
            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1 
            //写入标题             
            for (int i = 0; i < skinDataGridView1.ColumnCount; i++)
            { worksheet.Cells[1, i + 1] = skinDataGridView1.Columns[i].HeaderText; }
            //写入数值
            int rr = 0;
            for (int r = 0; r < skinDataGridView1.Rows.Count; r++)
            {

                if (skinDataGridView1.Rows[r].Visible)
                {

                    for (int i = 0; i < skinDataGridView1.ColumnCount; i++)
                    {
                        worksheet.Cells[rr + 2, i + 1] = skinDataGridView1.Rows[r].Cells[i].Value;
                    }
                    rr++;
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            worksheet.Columns.EntireColumn.AutoFit();//列宽自适应
            MessageBox.Show(fileName + "资料保存成功", "提示", MessageBoxButtons.OK);
            if (saveFileName != "")
            {
                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(saveFileName);  //fileSaved = true;                 
                }
                catch (Exception ex)
                {//fileSaved = false;                      
                    MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                }
            }
            xlApp.Quit();
            GC.Collect();//强行销毁  
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            refresh();
        }
      
          

















    }
}
