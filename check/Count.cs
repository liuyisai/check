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
using Microsoft.Office.Interop.Excel;

namespace check
{
    public partial class Count : Skin_Mac
    {
       System.Data.DataTable mainDt=null ;
       public Count(System.Data.DataTable dt)
        {
            InitializeComponent();
            mainDt = dt;
        }
        
        private static Count instance;

        public static Count CreateForm(System.Data.DataTable dt1) 
        {
            if (instance ==null||instance.IsDisposed)
            {
                instance = new Count(dt1);
                
            }
            return instance;
        }

       




















       








        private void skinButton1_Click(object sender, EventArgs e)
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
            GC.Collect();//强行销毁           }
        }

        private void Count_Load(object sender, EventArgs e)
        {
            #region  总体统计
            int totalNum = 0;
            int attendNum = 0;
            int unattendNum = 0;
            for (int i = 0; i <mainDt.Rows.Count; i++)
            {
                totalNum++;
                if (mainDt.Rows[i]["Column6"].ToString()=="是")
                {
                    attendNum++;

                }
                else
                    unattendNum++;

                
            }
            textBox1.Text = totalNum.ToString();
            textBox2.Text = attendNum.ToString();
            textBox3.Text = unattendNum.ToString();
            #endregion

            #region 具体信息
            int flag=0;
            int k = 0 ;
            int formalSum = 0, attendSum = 0, specialSum = 0, dueSum = 0, unarrivalSum = 0;
            string delegation="",delegationNext="";
            int formalerNum = 0,attenderNum=0,specialerNum=0,dueNum=0,unarriveNum=0;
            string delegator,attend;

            while (flag==0)
            {
                for (int i = k; i < mainDt.Rows.Count; i++)
                {
                    delegation = mainDt.Rows[i]["Column5"].ToString();
                    if (i+1<mainDt.Rows.Count)
                        delegationNext = mainDt.Rows[i + 1]["Column5"].ToString();
                    else
                        delegationNext = "";                  
                    delegator = mainDt.Rows[i]["Column2"].ToString();
                    attend = mainDt.Rows[i]["Column6"].ToString();
                    switch (delegator)
                    {
                        case "正式代表": 
                            {
                                if (attend=="否")
                                {
                                    formalerNum++;
                                    unarriveNum++;
                                    
                                }
                            }break;
                        case "列席代表":
                            {
                                if (attend == "否")
                                {
                                    attenderNum++;
                                    unarriveNum++;

                                }
                            }break;
                        case "特邀代表":
                            {
                                if (attend == "否")
                                {
                                    specialerNum++;
                                    unarriveNum++;

                                }
                            }break;
                       
                    }
                    dueNum++;
                    if (delegation==delegationNext)
                    {
                        continue;
                    }
                    else
                    {
                        k = i+1;
                        break;
                    }
                    
                }
                skinDataGridView1.Rows.Add(delegation,formalerNum,attenderNum,specialerNum,dueNum,unarriveNum);
                formalSum = formalSum + formalerNum;
                attendSum = attendSum + attenderNum;
                specialSum = specialSum + specialerNum;
                dueSum = dueSum + dueNum;
                unarrivalSum = unarrivalSum + unarriveNum;
                formalerNum = 0; attenderNum = 0; specialerNum = 0; dueNum = 0; unarriveNum = 0;
                if (k == mainDt.Rows.Count)
                    flag = 1;
                else
                    flag = 0;
            }

            skinDataGridView1.Rows.Add("总计",formalSum,attendSum,specialSum,dueSum,unarrivalSum);

	

            #endregion


           
          






        }
    }
}
