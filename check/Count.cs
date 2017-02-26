﻿using System;
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
        public Count()
        {
            InitializeComponent();
            skinDataGridView1.Rows.Add("1",1);
            skinDataGridView1.Rows.Add("2", 2);
            skinDataGridView1.Rows.Add("3", 3);
            skinDataGridView1.Rows[1].Visible = false;
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
    }
}
