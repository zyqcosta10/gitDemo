using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Collections;
using System.Data;
using System.Runtime.InteropServices;

namespace WindowsFormsApp1
{
    //class ExcelOperate
    //{

    //}
    public class ExcelOperate
    {
        //导入数据2003.11.28（增添了可设置是否表的第一列是自增长的id列）
        public bool InDataFromExcelColumns(OleDbConnection conn, string tableName, string[] tableColsName, string excelFilePath, string[] excelColsName, bool hasSysID)
        {
            if (hasSysID)
            {
                return this.InDataFromExcelColumns(conn, tableName, tableColsName, excelFilePath, excelColsName);
            }
            else
            {
                //把表的结构全部提取出来
                OleDbDataAdapter oleDA = new OleDbDataAdapter("select * from " + tableName, conn);
                DataSet ds = new DataSet();
                OleDbCommandBuilder oleCmdBlt = new OleDbCommandBuilder(oleDA);

                if (conn.State == ConnectionState.Closed)
                    conn.Open();
                oleDA.Fill(ds);
                conn.Close();

                System.Data.DataTable objDataTable = ds.Tables[0];
                if (!this.DataInTableColsFromExcelCols(objDataTable, tableColsName, excelFilePath, excelColsName, false))
                    return false;
                conn.Open();
                int addNum = oleDA.Update(ds);
                conn.Close();
                if (addNum != 0)
                    return true;
                return false;
            }
        }
        //把Excel中的数据导入具体的物理表中（单表）
        public bool InDataFromExcelColumns(OleDbConnection conn, string tableName, string[] tableColsName, string excelFilePath, string[] excelColsName)
        {
            // try
            // {
            //把表的结构全部提取出来
            OleDbDataAdapter oleDA = new OleDbDataAdapter("select * from " + tableName, conn);
            DataSet ds = new DataSet();
            OleDbCommandBuilder oleCmdBlt = new OleDbCommandBuilder(oleDA);

            if (conn.State == ConnectionState.Closed)
                conn.Open();
            oleDA.Fill(ds);
            conn.Close();

            System.Data.DataTable objDataTable = ds.Tables[0];
            if (!this.DataInTableColsFromExcelCols(objDataTable, tableColsName, excelFilePath, excelColsName))
                return false;
            conn.Open();
            int addNum = oleDA.Update(ds);
            conn.Close();
            if (addNum != 0)
                return true;
            return false;
            // }
            // catch
            // {
            // return false ;
            // }
        }
        //由Excel中的数据完全导入到表中（前提是excel的表结构与要导入的表的结构是完全一样的）
        public bool DataInTableFromExcel(System.Data.DataTable objDataTable, string excelFilePath)
        {
            //先得到excel文件对应的表
            System.Data.DataTable excelDataTable = this.GetDataFromExcel(excelFilePath);
            if (excelDataTable == null || excelDataTable.Rows.Count == 0
            || ((excelDataTable.Rows.Count == 1) && (excelDataTable.Columns.Count == 1) && (excelDataTable.Rows[0][0].ToString().Equals(""))))
                return false;
            if (objDataTable == null)
                return false;
            if (excelDataTable.Columns.Count != objDataTable.Columns.Count)
                return false;
            //检测数据类型是否匹配
            for (int i = 0; i < objDataTable.Columns.Count; i++)
            {
                if (!this.IsMatchDataType(objDataTable.Columns[i].DataType.ToString(), excelDataTable.Columns[i].DataType.ToString()))
                    return false;
            }
            //完全添加进去
            System.Data.DataRow myDataRow;

            for (int i = 0; i < excelDataTable.Rows.Count; i++)
            {
                myDataRow = objDataTable.NewRow();
                for (int j = 0; j < objDataTable.Columns.Count; j++)
                {
                    myDataRow[j] = this.GetMatchData(objDataTable.Columns[j].DataType.ToString(), excelDataTable.Rows[i][j].ToString().Trim());
                }

                objDataTable.Rows.Add(myDataRow);
            }

            return true;
        }

        /// <summary>
        /// 给指定的表的某列更新成Excel中某列的数据
        /// </summary>
        /// <param name="conn">对此表有效的连接</param>
        /// <param name="tableName">物理表的表名</param>
        /// <param name="PID">此物理表的主键列名称（这里认为主键列只有一个）</param>
        /// <param name="tableColsName">此物理表中要更新的列集合</param>
        /// <param name="srcDataTable">更新的数据源</param>
        /// <param name="srcColsName">对应tableColsName中数据源中的列名集合</param>
        /// <returns></returns>
        public bool DataUpdateTableColsFromExcelCols(OleDbConnection conn, string tableName, string ID, string[] tableColsName, System.Data.DataTable srcDataTable, string[] srcColsName)
        {
            //此时认为objDataTable与srcDataTable的第一列的id列是连接关系
            OleDbDataAdapter oleDA = new OleDbDataAdapter("select * from " + tableName, conn);
            DataSet ds = new DataSet();
            OleDbCommandBuilder oleCmdBlt = new OleDbCommandBuilder(oleDA);

            if (conn.State == ConnectionState.Closed)
                conn.Open();
            oleDA.Fill(ds);
            conn.Close();

            System.Data.DataTable objDataTable = ds.Tables[0];
            //设主键
            DataColumn[] keyCols = new DataColumn[1];
            keyCols[0] = objDataTable.Columns[ID];
            objDataTable.PrimaryKey = keyCols;

            //循环检测物理表的每一条记录是不是要更新
            for (int i = 0; i < objDataTable.Rows.Count; i++)
            {
                string idValue = objDataTable.Rows[i][ID].ToString().Trim();
                for (int j = 0; j < srcDataTable.Rows.Count; j++)
                {
                    string tempID = srcDataTable.Rows[j][0].ToString().Trim();
                    if (tempID.Equals(idValue))
                    {
                        objDataTable.Rows[i].BeginEdit();
                        //=================================
                        for (int k = 0; k < objDataTable.Columns.Count; k++)
                        {
                            int index = this.MatchStringIndex(objDataTable.Columns[k].ColumnName, tableColsName);
                            if (index != -1)
                            {
                                objDataTable.Rows[i][tableColsName[index]] = this.GetMatchData(objDataTable.Columns[k].DataType.ToString().Trim(), srcDataTable.Rows[j][srcColsName[index]].ToString().Trim());
                            }
                        }
                        //=================================
                        objDataTable.Rows[i].EndEdit();

                        break;
                    }
                }
            }//for

            conn.Open();
            int affect = oleDA.Update(ds);
            conn.Close();
            if (affect == 0)
                return false;
            return true;
        }
        //2003.11.28 To （增添了可设置是否表的第一列是自增长的id列）
        public bool DataInTableColsFromExcelCols(System.Data.DataTable objDataTable, string[] tableColsName, string excelFilePath, string[] excelColsName, bool hasID)
        {
            //此时认为objDataTable是都有自动增长的id列
            if (tableColsName == null || excelColsName == null || tableColsName.Length != excelColsName.Length)
                return false;

            //先得到excel文件对应的表
            System.Data.DataTable excelDataTable = this.GetDataFromExcel(excelFilePath);

            if (excelDataTable == null || excelDataTable.Rows.Count == 0
            || ((excelDataTable.Rows.Count == 1) && (excelDataTable.Columns.Count == 1) && (excelDataTable.Rows[0][0].ToString().Equals(""))))
                return false;
            if (objDataTable == null)
                return false;
            //把相应的数据添加到目标表中
            //首先验证要添加的数据与源数据的类型是否匹配
            // for( int i = 0 ; i < excelColsName.Length ; i++ )
            // {
            // string srcDataType = excelDataTable.Columns[ excelColsName[i] ].DataType.ToString( ).Trim( ) ;
            // string objDataType = objDataTable.Columns[ tableColsName[i] ].DataType.ToString( ).Trim( ) ;
            //
            // if( ! this.IsMatchDataType( objDataType , srcDataType ) )
            // return false ;
            // }

            //System.Data.DataRow myDataRow ;

            //添加（一次添加一行）

            //myDataRow = objDataTable.NewRow( ) ;
            int colNum = objDataTable.Columns.Count;
            //object [] colValues = new Object[ colNum ] ;

            int dataRow = excelDataTable.Rows.Count;
            for (int k = 0; k < dataRow; k++)
            {
                //int colNum = objDataTable.Columns.Count ;
                object[] colValues = new Object[colNum];

                //认为第一列是自动增长的id列,不添加任何值，由系统自动生成
                //colValues[0] = Guid.NewGuid( ) ;

                int addIndex = 0;
                int j = 0;
                if (hasID) j = 1;
                for (; j < colNum; j++)
                {
                    if (this.NumberInSource(objDataTable.Columns[j].ColumnName, tableColsName) != 1)
                    {
                        colValues[j] = this.GetMatchData(objDataTable.Columns[j].DataType.ToString(), "");
                        //myDataRow[j] = this.GetMatchData( objDataTable.Columns[j].DataType.ToString( ) , "") ;
                    }
                    else//是要添加数据的列时
                    {
                        //找到与此列对应的tableColsName中的项索引
                        int index = MatchStringIndex(objDataTable.Columns[j].ColumnName, tableColsName);
                        colValues[j] = this.GetMatchData(objDataTable.Columns[j].DataType.ToString(), excelDataTable.Rows[k][excelColsName[index]].ToString().Trim());
                        //myDataRow[j] = this.GetMatchData( objDataTable.Columns[j].DataType.ToString( ) ,excelDataTable.Rows[k][excelColsName[addIndex++]].ToString().Trim() );
                    }
                }

                objDataTable.Rows.Add(colValues);
            }

            return true;
        }
        //给指定的表的某列导入Excel中某列的数据
        public bool DataInTableColsFromExcelCols(System.Data.DataTable objDataTable, string[] tableColsName, string excelFilePath, string[] excelColsName)
        {
            //此时认为objDataTable是都有自动增长的id列
            if (tableColsName == null || excelColsName == null || tableColsName.Length != excelColsName.Length)
                return false;

            //先得到excel文件对应的表
            System.Data.DataTable excelDataTable = this.GetDataFromExcel(excelFilePath);

            if (excelDataTable == null || excelDataTable.Rows.Count == 0
            || ((excelDataTable.Rows.Count == 1) && (excelDataTable.Columns.Count == 1) && (excelDataTable.Rows[0][0].ToString().Equals(""))))
                return false;
            if (objDataTable == null)
                return false;
            //把相应的数据添加到目标表中
            //首先验证要添加的数据与源数据的类型是否匹配
            // for( int i = 0 ; i < excelColsName.Length ; i++ )
            // {
            // string srcDataType = excelDataTable.Columns[ excelColsName[i] ].DataType.ToString( ).Trim( ) ;
            // string objDataType = objDataTable.Columns[ tableColsName[i] ].DataType.ToString( ).Trim( ) ;
            //
            // if( ! this.IsMatchDataType( objDataType , srcDataType ) )
            // return false ;
            // }

            //System.Data.DataRow myDataRow ;

            //添加（一次添加一行）

            //myDataRow = objDataTable.NewRow( ) ;
            int colNum = objDataTable.Columns.Count;
            //object [] colValues = new Object[ colNum ] ;

            int dataRow = excelDataTable.Rows.Count;
            for (int k = 0; k < dataRow; k++)
            {
                //int colNum = objDataTable.Columns.Count ;
                object[] colValues = new Object[colNum];

                //认为第一列是自动增长的id列,不添加任何值，由系统自动生成
                //colValues[0] = Guid.NewGuid( ) ;

                int addIndex = 0;

                for (int j = 1; j < colNum; j++)
                {
                    if (this.NumberInSource(objDataTable.Columns[j].ColumnName, tableColsName) != 1)
                    {
                        colValues[j] = this.GetMatchData(objDataTable.Columns[j].DataType.ToString(), "");
                        //myDataRow[j] = this.GetMatchData( objDataTable.Columns[j].DataType.ToString( ) , "") ;
                    }
                    else//是要添加数据的列时
                    {
                        //找到与此列对应的tableColsName中的项索引
                        int index = MatchStringIndex(objDataTable.Columns[j].ColumnName, tableColsName);
                        colValues[j] = this.GetMatchData(objDataTable.Columns[j].DataType.ToString(), excelDataTable.Rows[k][excelColsName[index]].ToString().Trim());
                        //myDataRow[j] = this.GetMatchData( objDataTable.Columns[j].DataType.ToString( ) ,excelDataTable.Rows[k][excelColsName[addIndex++]].ToString().Trim() );
                    }
                }

                objDataTable.Rows.Add(colValues);
            }

            return true;
        }
        //把数据表的内容导出到Excel文件中:method1
        public bool OutDataToExcel(System.Data.DataTable srcDataTable, string excelFilePath)
        {
            if (srcDataTable == null)
                return false;
            Excel.Application myExcel = new Excel.Application();
            try
            {
                object oMissing = System.Reflection.Missing.Value;

                myExcel.Application.Workbooks.Open(excelFilePath, oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing,oMissing, oMissing);

                Excel._Workbook xBk;
                Excel._Worksheet xSt;

                xBk = myExcel.Workbooks.Add(true);
                xSt = (Excel._Worksheet)xBk.ActiveSheet;


                int rowIndex = 1;
                int colIndex = 1;
                //写标题
                foreach (DataColumn col in srcDataTable.Columns)
                {
                    xSt.Cells[rowIndex, colIndex] = col.ColumnName.Trim();
                    ++colIndex;
                }
                rowIndex = 2;
                foreach (DataRow row in srcDataTable.Rows)
                {
                    colIndex = 1;
                    foreach (DataColumn col in srcDataTable.Columns)
                    {
                        xSt.Cells[rowIndex, colIndex] = row[col.ColumnName].ToString().Trim();
                        ++colIndex;
                    }
                    ++rowIndex;
                }
                Marshal.ReleaseComObject(myExcel);
                return true;
            }
            catch
            {
                Marshal.ReleaseComObject(myExcel);
                return false;
            }
        }
        //从Excel中的workBook[1]中的全部的sheet选取全部的数据
        public System.Data.DataTable GetDataFromExcel(string excelFilePath)
        {
            Excel.Application myExcel = new Excel.Application();
            try
            {
                //取得Excel文件中共有的sheet的数目

                object oMissing = System.Reflection.Missing.Value;

                myExcel.Application.Workbooks.Open(excelFilePath, oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                Excel.Workbook myBook = myExcel.Workbooks[1];
                int sheetNum = myBook.Worksheets.Count;
                //===========2003.11.28=============
                string[] sheetName = new String[sheetNum];
                int sheetIndex = 0;
                foreach (Excel.Worksheet xlsheet in myBook.Worksheets)
                {
                    sheetName[sheetIndex++] = xlsheet.Name;
                }
                //===========#2003.11.28============

                myExcel.Application.Workbooks.Close();

                string strConn = String.Format(" Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = {0};Extended Properties=Excel 8.0", excelFilePath);

                OleDbConnection conn = new OleDbConnection(strConn);

                string selSqlStr = "";

                //循环取完所有的sheet中的数据
                //
                System.Data.DataTable totalDataTable = new System.Data.DataTable();

                for (int i = 1; i <= 3; i++)
                {
                    string sheetFlag = sheetName[i - 1];
                    if (sheetFlag.Contains("第"))
                    {
                        selSqlStr = "select * from [" + sheetName[i - 1] + "$]";
                        //selSqlStr = "select * from [Sheet3$]" ;//==================right
                        OleDbDataAdapter oleDa = new OleDbDataAdapter(selSqlStr, conn);

                        DataSet ds = new DataSet();
                        conn.Open();
                        oleDa.Fill(ds, "dataTable");
                        conn.Close();

                        if (ds.Tables.Count != 0)
                        {
                            totalDataTable = this.AddDataTable(totalDataTable, ds.Tables["dataTable"]);
                        }
                    }

                  
                }
                Marshal.ReleaseComObject(myExcel);
                return totalDataTable;
            }
            catch
            {
                Marshal.ReleaseComObject(myExcel);
                return null;
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <param name="index">sheet的索引号,从1开始</param>
        /// <returns></returns>
        public string[] GetColumnsNameList(string excelFilePath, int index)
        {

            //string [] returnList = new string[ dataTable.Columns.Count ] ;
            Excel.Application myExcel = new Excel.Application();
            try
            {
                //取得Excel文件中共有的sheet的数目

                object oMissing = System.Reflection.Missing.Value;

                myExcel.Application.Workbooks.Open(excelFilePath, oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                Excel.Workbook myBook = myExcel.Workbooks[1];
                Excel.Sheets sheets = myBook.Worksheets;
                Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(index);
                Excel.Range range = worksheet.get_Range("A1", "CZ1");
                System.Array myvalues = (System.Array)range.Cells.Value2;

                System.Collections.ArrayList array = this.ConvertToStringArray(myvalues);
                //int colName = myBook.Worksheets.Count ;
                //===========2003.11.28=============
                // string [] sheetName = new String[ sheetNum ] ;
                // int sheetIndex = 0 ;
                // foreach( Excel.Worksheet xlsheet in myBook.Worksheets )
                // {
                // sheetName[sheetIndex++] = xlsheet.Name ;
                // }
                string[] colName = new string[array.Count];
                //myExcel.get_Range(
                for (int i = 0; i < array.Count; i++)
                {
                    colName[i] = array[i].ToString().Trim();
                }
                myExcel.Application.Workbooks.Close();

                Marshal.ReleaseComObject(myExcel);
                return colName;
            }
            catch
            {
                Marshal.ReleaseComObject(myExcel);
                return null;
            }
        }

        /// <summary>
        /// 获取有指定的colList的页的名称
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        public string[] GetSheetNameList(string excelFilePath, string[] colList)
        {
            Excel.Application myExcel = new Excel.Application();
            try
            {
                //取得Excel文件中共有的sheet的数目

                object oMissing = System.Reflection.Missing.Value;

                myExcel.Application.Workbooks.Open(excelFilePath, oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                Excel.Workbook myBook = myExcel.Workbooks[1];
                int sheetNum = myBook.Worksheets.Count;
                string[] col = null;
                ArrayList totalName = new ArrayList();
                for (int i = 1; i <= sheetNum; i++)
                {
                    col = this.GetColumnsNameList(excelFilePath, i);
                    if (col == null || col.Length == 0 || col.Length != colList.Length)
                        continue;
                    for (int j = 0; j < col.Length; j++)
                    {
                        if (col[j].Trim() != colList[j].Trim())
                            continue;
                    }
                    Excel.Worksheet xlsheet = (Excel.Worksheet)myBook.Worksheets[i];
                    totalName.Add(xlsheet.Name);
                }
                //===========2003.11.28=============
                string[] sheetName = new String[totalName.Count];

                for (int i = 0; i < totalName.Count; i++)
                {
                    sheetName[i] = totalName[i].ToString().Trim();
                }
                myExcel.Application.Workbooks.Close();

                Marshal.ReleaseComObject(myExcel);
                return sheetName;
            }
            catch
            {
                Marshal.ReleaseComObject(myExcel);
                return null;
            }
        }

        /// <summary>
        /// 获取有指定的colList的页的名称
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        public string[] GetSheetNameList(string excelFilePath)
        {
            Excel.Application myExcel = new Excel.Application();
            try
            {
                //取得Excel文件中共有的sheet的数目

                object oMissing = System.Reflection.Missing.Value;

                myExcel.Application.Workbooks.Open(excelFilePath, oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                Excel.Workbook myBook = myExcel.Workbooks[1];
                int sheetNum = myBook.Worksheets.Count;
                string[] col = null;
                ArrayList totalName = new ArrayList();
                for (int i = 1; i <= sheetNum; i++)
                {
                    col = this.GetColumnsNameList(excelFilePath, i);
                    if (col == null || col.Length == 0)
                        continue;
                    //for (int j = 0; j < col.Length; j++)
                    //{
                    //    if (col[j].Trim() != colList[j].Trim())
                    //        continue;
                    //}
                    Excel.Worksheet xlsheet = (Excel.Worksheet)myBook.Worksheets[i];
                    totalName.Add(xlsheet.Name);
                }
                //===========2003.11.28=============
                string[] sheetName = new String[totalName.Count];

                for (int i = 0; i < totalName.Count; i++)
                {
                    sheetName[i] = totalName[i].ToString().Trim();
                }
                myExcel.Application.Workbooks.Close();

                Marshal.ReleaseComObject(myExcel);
                return sheetName;
                myExcel.Application.Workbooks.Close();
            }
            catch
            {
                Marshal.ReleaseComObject(myExcel);
                return null;
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        public bool DelSheetColumnData(string excelFilePath, System.Collections.ArrayList colName)
        {
            Excel.Application myExcel = new Excel.Application();
            try
            {
                if (colName == null && colName.Count == 0)
                    return true;
                //取得Excel文件中共有的sheet的数目

                object oMissing = System.Reflection.Missing.Value;

                myExcel.Application.Workbooks.Open(excelFilePath, oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                Excel.Workbook myBook = myExcel.Workbooks[1];
                Excel.Sheets sheets = myBook.Worksheets;
                Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);
                Excel.Range range = worksheet.get_Range("A1", "Z1");
                System.Array myvalues = (System.Array)range.Cells.Value2;

                Excel.Range delRange;
                for (int i = 1; i <= myvalues.Length; i++)
                {
                    if (ISDelCol(myvalues.GetValue(1, i).ToString().Trim(), colName) != -1)
                    {
                        delRange = worksheet.get_Range(i, i);
                        delRange.Delete(i);
                    }
                    // if (values.GetValue(1, i) != null && values.GetValue(1,i).ToString().Trim() != "" )
                    // theArray.Add( values.GetValue( 1, i ).ToString().Trim() );
                }

                myExcel.Application.Workbooks.Close();

                Marshal.ReleaseComObject(myExcel);
                return true;
            }
            catch
            {
                Marshal.ReleaseComObject(myExcel);
                return false;
            }
        }
        /// <summary>
        ///
        /// </summary>
        /// <param name="oldColName"></param>
        /// <param name="newColName"></param>
        /// <returns></returns>
        public void SetColumnName(string excelFilePath, ArrayList oldColName, ArrayList newColName)
        {
            Excel.Application myExcel = new Excel.Application();
            try
            {
                if (oldColName == null || oldColName.Count == 0)
                    return;
                //取得Excel文件中共有的sheet的数目

                object oMissing = System.Reflection.Missing.Value;

                myExcel.Application.Workbooks.Open(excelFilePath, oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                Excel.Workbook myBook = myExcel.Workbooks[1];
                Excel.Sheets sheets = myBook.Worksheets;
                Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);
                Excel.Range range = worksheet.get_Range("A1", "Z1");
                System.Array myvalues = (System.Array)range.Cells.Value2;

                Excel._Workbook xBk;
                Excel._Worksheet xSt;

                xBk = myExcel.Workbooks[1];
                xSt = (Excel._Worksheet)xBk.ActiveSheet;

                //Excel.Range delRange ;
                for (int i = 1; i <= myvalues.Length; i++)
                {
                    int index = ISDelCol(myvalues.GetValue(1, i).ToString().Trim(), oldColName);
                    if (index != -1)
                    {
                        xSt.Cells[1, i] = newColName[i].ToString().Trim();
                    }
                }

                myExcel.Application.Workbooks.Close();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);

            }
            catch
            {
                Marshal.ReleaseComObject(myExcel);

            }
        }
        /// <summary>
        ///
        /// </summary>
        public void ReleaseComExcel()
        {
        }

        /// <summary>
        /// Release
        /// </summary>
        //public void Dispose(string FullName)
        //{
        //    tworkbook.Close(false, FullName, missing);
        //    app.Workbooks.Close();
        //    app.Quit();
        //    if (app != null)
        //    {
        //        foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName("Excel"))
        //        {
        //            //先判断当前进程是否是excel  
        //            if (!p.CloseMainWindow())
        //            {
        //                p.Kill();
        //            }
        //        }
        //    }
        //    tworkbook = null;
        //    app = null;
        //    GC.Collect();
        //}


        #region 内部接口
        //在表中追加与此表结构相同的表中的记录
        System.Data.DataTable AddDataTable(System.Data.DataTable oldDataTable, System.Data.DataTable srcDataTable)
        {

            if (srcDataTable == null)
                return oldDataTable;

            System.Data.DataTable returnDataTable = new System.Data.DataTable();
            System.Data.DataRow myDataRow;

            //
            if (oldDataTable != null)
            {
                if (oldDataTable.Columns.Count != srcDataTable.Columns.Count)
                {
                    if (oldDataTable.Columns.Count >= srcDataTable.Columns.Count)
                        return oldDataTable;
                    return srcDataTable;
                }
                //给表格添加列
                for (int i = 0; i < srcDataTable.Columns.Count; i++)
                {
                    returnDataTable.Columns.Add(new System.Data.DataColumn(srcDataTable.Columns[i].ColumnName, srcDataTable.Columns[i].DataType));
                }

                int rowNum = oldDataTable.Rows.Count;

                //添加每一行的数据到目标数据表中
                for (int i = 0; i < rowNum; i++)
                {
                    myDataRow = returnDataTable.NewRow();
                    int colNum = oldDataTable.Columns.Count;

                    for (int j = 0; j < colNum; j++)
                    {
                        myDataRow[j] = oldDataTable.Rows[i].ItemArray[j];
                    }

                    returnDataTable.Rows.Add(myDataRow);
                }

                rowNum = srcDataTable.Rows.Count;

                for (int i = 0; i < rowNum; i++)
                {
                    myDataRow = returnDataTable.NewRow();
                    int colNum = srcDataTable.Columns.Count;
                    if (colNum > 0)
                    {
                        if (((srcDataTable.Rows[i].ItemArray[0]).ToString() == null
                        || (srcDataTable.Rows[i].ItemArray[0]).ToString().Equals(""))
                        && (rowNum == 1) && (colNum == 1))
                            return oldDataTable;

                        for (int j = 0; j < colNum; j++)
                        {
                            myDataRow[j] = this.GetMatchData(srcDataTable.Columns[j].DataType.ToString(), srcDataTable.Rows[i][j].ToString());
                        }

                        returnDataTable.Rows.Add(myDataRow);
                    }
                }
            }

            return returnDataTable;
        }
        //取得对应数据类型的数据
        object GetMatchData(string dataType, string data)
        {
            switch (dataType)
            {
                case "System.Double":
                    if (data.Equals(""))
                        return 0;
                    return Convert.ToDouble(data);

                case "System.Decimal":
                    if (data.Equals(""))
                        return 0;
                    return Convert.ToDecimal(data);

                case "System.DateTime":
                    if (data.Equals(""))
                    {
                        return Convert.ToDateTime(System.DateTime.Now.ToLongDateString());
                    }
                    return Convert.ToDateTime(data);

                case "System.Int64":
                    if (data.Equals(""))
                        return 0;
                    return Convert.ToInt64(data);

                case "System.Int32":
                    if (data.Equals(""))
                        return 0;
                    return Convert.ToInt32(data);

                case "System.Int16":
                    if (data.Equals(""))
                        return 0;
                    return Convert.ToInt16(data);

                case "System.Boolean":
                    if (data.Equals(""))
                        return 0;
                    return Convert.ToInt16(data);
            }

            return data;
        }

        //验证数据类型的匹配
        public bool IsMatchDataType(string objDataType, string srcDataType)
        {
            if (objDataType.Equals(srcDataType))
                return true;
            if ((!objDataType.Equals("System.String")) && (!objDataType.Equals("System.DateTime")))
            {
                switch (srcDataType)
                {
                    case "System.Int32":
                    case "System.Int16":
                    case "System.Double":
                    case "System.Decimal":
                        return true;
                }
                return false;
            }

            if (objDataType.Equals("System.DateTime"))
            {
                switch (srcDataType)
                {
                    case "System.DateTime":
                    case "System.String":
                        return true;
                }

                return false;
            }

            return false;
        }
        //获取在字符串集合中与目标字符串值相同的字符串的个数
        private int NumberInSource(string str, string[] source)
        {
            if (source == null)
                return 0;
            if (str == null)
                return 0;
            int num = 0;
            for (int i = 0; i < source.Length; i++)
            {
                if (str.Equals(source[i]))
                    num++;
            }
            return num;
        }
        private int MatchStringIndex(string objStr, string[] srcStr)
        {
            if (srcStr != null)
            {
                for (int i = 0; i < srcStr.Length; i++)
                {
                    if (objStr.Trim().Equals(srcStr[i].Trim()))
                        return i;
                }
            }

            return -1;
        }
        System.Collections.ArrayList ConvertToStringArray(System.Array values)
        {
            System.Collections.ArrayList theArray = new System.Collections.ArrayList();
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(1, i) != null && values.GetValue(1, i).ToString().Trim() != "")
                    theArray.Add(values.GetValue(1, i).ToString().Trim());
            }
            return theArray;
        }
        int ISDelCol(string objColName, System.Collections.ArrayList colName)
        {
            if (colName == null) return -1;
            for (int i = 0; i < colName.Count; i++)
            {
                if (objColName.Trim() == colName[i].ToString().Trim())
                    return i;
            }
            return -1;
        }
        #endregion
    }


}
