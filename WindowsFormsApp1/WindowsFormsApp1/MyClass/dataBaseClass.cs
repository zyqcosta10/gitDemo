using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;

namespace WindowsFormsApp1.MyClass
{
    class dataBaseClass
    {
        /// <summary>
        /// 连接数据库公共方法
        /// </summary>
        /// <returns>数据库连接</returns>
        public static SqlConnection DBCon()
        {
            string connString = "server=127.0.0.1;Initial Catalog=TravelAgency;User ID=sa;pwd=yanxin@502";
            return new SqlConnection(connString);
        }

        /// <summary>
        /// 执行SQL语句(返回执行成功的行数)
        /// </summary>
        /// <param name="strSql">操作数据库的SQL语句</param>
        /// <returns>返回执行成功的行数</returns>
        public static int OperateData(string strSql)
        {
            SqlConnection conn = DBCon();
            conn.Open();
            SqlCommand cmd = new SqlCommand(strSql, conn);
            int i = 0;
            try
            {
                i = cmd.ExecuteNonQuery();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString());
                i = 0;
            }
            conn.Close();
            return i;
        }



        /// <summary>
        /// 执行SQL读取语句
        /// </summary>
        /// <param name="strsql1">操作数据的语句</param>
        /// <returns>数据读取集合</returns>
        public static SqlDataReader OperateDataReader(string strsql1)
        {
            SqlConnection conn = DBCon();
            conn.Open();
            SqlCommand cmd = new SqlCommand(strsql1, conn);
            SqlDataReader sdr = cmd.ExecuteReader();
            sdr.Read();
            sdr.Close();
            return sdr;          
        }

        /// <summary>
        /// 执行SQL查询语句(dataset)
        /// </summary>
        /// <param name="strsqlset">查询语句</param>
        /// <returns>dataset</returns>
        public static DataSet OperateDataSet(string strsqlset)
        {

            try
            {
                SqlConnection conn = DBCon();
                conn.Open();
                SqlDataAdapter sda = new SqlDataAdapter(strsqlset, conn);
                DataSet ds = new DataSet();
                sda.Fill(ds);
                conn.Close();
                return ds;
            }
            catch
            {
                return null;
            }
        }


    }
}
