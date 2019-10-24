
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using DfkjToolKit;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        DataTable GetTB = null;
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //DataGet();
        }

        private void DataGet()
        {
            //string sql = "select * from CheckWork";
            //dataGridView1.DataSource = MyClass.dataBaseClass.OperateDataSet(sql).Tables[0].DefaultView;
        }

        /// <summary>
        /// 读取Excel文件（孔隙度、渗透率）
        /// </summary>
        /// <param name="FilePath"></param>
        private void ReadExcel(string FilePath)
        {
            // ExcelOperate EO = new ExcelOperate();
            //Excel.Application myExcel = new Excel.Application();

            ////取得Excel文件中共有的sheet的数目

            //object oMissing = System.Reflection.Missing.Value;

            //myExcel.Application.Workbooks.Open(FilePath, oMissing, oMissing, oMissing, oMissing, oMissing,
            //oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //Excel.Workbook myBook = myExcel.Workbooks[1];
            //int sheetNum = myBook.Worksheets.Count;

            ExcelOpreate ExcelOp = new ExcelOpreate(FilePath);
            DataTable tblDatas = new DataTable("Datas");
            DataColumn dc = null;
            dc = tblDatas.Columns.Add("姓名", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("部门", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("日期", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("状态", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("考勤代码", Type.GetType("System.String"));
            //for (int j = 1; j < sheetNum; j++)
            //{
            ExcelOp.SheelTableIndex = 1;
            DataRow newRow;
            //string falg = ExcelOp.ReadExcelCellString(1, 2);
            //if (falg.Contains("月度汇总"))
            //{
            for (int i = 5; i < 200; i++)
            {
                progressBar1.Maximum = 100;
                if (ExcelOp.ReadExcelCellString(i, 1) != "")
                {
                    for (int j = 23; j < 54; j++)
                    {
                        newRow = tblDatas.NewRow();
                        newRow["姓名"] = ExcelOp.ReadExcelCellString(i, 1);
                        newRow["部门"] = ExcelOp.ReadExcelCellString(i, 2);
                        newRow["日期"] = ExcelOp.ReadExcelCellString(3, j);
                        newRow["状态"] = ExcelOp.ReadExcelCellString(i, j);
                        newRow["考勤代码"] = GetWorkCode(ExcelOp.ReadExcelCellString(i, j));
                        //newRow["渗透率水平"] = ExcelOp.ReadExcelCellString(i, 8);
                        //newRow["渗透率垂直"] = ExcelOp.ReadExcelCellString(i, 9);
                        tblDatas.Rows.Add(newRow);
                    }
                    progressBar1.Value = i + 1;
                }
                else
                {
                    break;
                }
                progressBar1.Value = 100;
            }
            //}
            //progressBar1.Maximum = sheetNum;
            //progressBar1.Value = j + 1;
            //}           
            dataGridView1.DataSource = tblDatas;
            GetTB = tblDatas;
            //MessageBox.Show("共" + tblDatas.Rows.Count.ToString() + "行");
            ExcelOp.KillExcel();
        }


        private string GetWorkCode(string WorkState)
        {
            if (WorkState == "正常")
            {
                return "P";
            }
            else if (WorkState == "休息")
            {
                return "O";
            }
            else if (WorkState == "病假")
            {
                return "S";
            }
            else if (WorkState == "年假")
            {
                return "A";
            }
            else if (WorkState == "旷工")
            {
                return "X";
            }
            else if (WorkState == "事假")
            {
                return "E";
            }
            else if (WorkState == "早退")
            {
                return "LF";
            }
            else if (WorkState == "辞职")
            {
                return "R";
            }
            else if (WorkState == "加班")
            {
                return "OT";
            }
            else if (WorkState == "无薪年假")
            {
                return "UP";
            }
            else if (WorkState == "补假")
            {
                return "L";
            }
            else if (WorkState == "迟到")
            {
                return "LR";
            }
            else if (WorkState == "产假")
            {
                return "Y";
            }
            else if (WorkState == "丧假")
            {
                return "F";
            }
            else if (WorkState == "公出")
            {
                return "B";
            }
            else if (WorkState == "婚假")
            {
                return "M";
            }
            else if (WorkState == "法定假日")
            {
                return "PH";
            }
            else
            {
                return WorkState;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ReadExcel("E:\\TEST\\1.xlsx");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelOpreate ExcelOp = new ExcelOpreate("E:\\TEST\\2.xlsx");
            int count = 0;
            progressBar2.Maximum = 50;
            for (int i = 7; i < 45; i++)
            {
                ExcelOp.WriteExcelCell(i, 2, GetTB.Rows[count]["姓名"].ToString());
                for (int j = 0; j < 31; j++)
                {
                    ExcelOp.WriteExcelCell(i, j + 3, GetTB.Rows[count]["考勤代码"].ToString());
                    count++;
                }
                progressBar2.Value = i + 1;
            }
            progressBar2.Value = 50;
            ExcelOp.Save();
            ExcelOp.KillExcel();

        }
    }
}
