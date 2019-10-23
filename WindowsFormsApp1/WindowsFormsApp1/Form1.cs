
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
            DataGet();
        }

        private void DataGet()
        {
            string sql = "select * from CheckWork";
            dataGridView1.DataSource = MyClass.dataBaseClass.OperateDataSet(sql).Tables[0].DefaultView;
        }

        /// <summary>
        /// 读取Excel文件（孔隙度、渗透率）
        /// </summary>
        /// <param name="FilePath"></param>
        private void ReadExcel(string FilePath)
        {
            // ExcelOperate EO = new ExcelOperate();
            Excel.Application myExcel = new Excel.Application();

            //取得Excel文件中共有的sheet的数目

            object oMissing = System.Reflection.Missing.Value;

            myExcel.Application.Workbooks.Open(FilePath, oMissing, oMissing, oMissing, oMissing, oMissing,
            oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            Excel.Workbook myBook = myExcel.Workbooks[1];
            //int sheetNum = myBook.Worksheets.Count;

            ExcelOpreate ExcelOp = new ExcelOpreate(FilePath);
            DataTable tblDatas = new DataTable("Datas");
            DataColumn dc = null;



            dc = tblDatas.Columns.Add("样品号", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("井深", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("空隙度", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("渗透率水平", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("渗透率垂直", Type.GetType("System.String"));
            //for (int j = 1; j < sheetNum; j++)
            //{
            ExcelOp.SheelTableIndex = 1;
            DataRow newRow;
            //string falg = ExcelOp.ReadExcelCellString(1, 2);
            //if (falg.Contains("月度汇总"))
            //{
            for (int i = 5; i < 200; i++)
            {
                //if (!ExcelOp.ReadExcelCellString(i, 2).Contains("分析人") && ExcelOp.ReadExcelCellString(i, 2) != "")
                //{
                newRow = tblDatas.NewRow();
                newRow["样品号"] = ExcelOp.ReadExcelCellString(i, 1);
                newRow["井深"] = ExcelOp.ReadExcelCellString(i, 2);
                newRow["空隙度"] = ExcelOp.ReadExcelCellString(i, 7);
                newRow["渗透率水平"] = ExcelOp.ReadExcelCellString(i, 8);
                newRow["渗透率垂直"] = ExcelOp.ReadExcelCellString(i, 9);

                tblDatas.Rows.Add(newRow);
                //}
                //else
                //{
                //    break;
                //}
            }
            //}
            //progressBar1.Maximum = sheetNum;
            //progressBar1.Value = j + 1;
            //}

            dataGridView1.DataSource = tblDatas;
            GetTB = tblDatas;

            MessageBox.Show("共" + tblDatas.Rows.Count.ToString() + "行");
            ExcelOp.KillExcel();
        }




        private void button1_Click(object sender, EventArgs e)
        {
            ReadExcel("E:\\TEST\\1.xlsx");
        }
    }
}
