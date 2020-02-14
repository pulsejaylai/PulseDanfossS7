using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Data.SqlClient;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.HSSF.Util;
using System.IO;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for Expert.xaml
    /// </summary>
    public partial class Expert : Window
    {
        public Expert()
        {
            InitializeComponent();
        }
        private string modelname="";
        public static int dataitem;
        string tablename;
        SqlConnection con = new SqlConnection();
        SqlCommand cmd = new SqlCommand();
        DataTable dt = new DataTable();
        public string modelset
        {
            set
            {
                modelname = value;
            }
        }
        public int itemSet
        {
            set
            {
                dataitem = value;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            con.ConnectionString = "server=WIN-20180329ZKY\\SQLEXPRESS;database=Danfossdata;uid=sa;pwd=sqlte";
             tablename = "D" + modelname.Replace("-", "");
            try
            {
                con.Open();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
            //  string sqlcomm = "select Model,password from Admin2";
            //string sqlcomm = "select *from [Danfossdata].[dbo].[Testdata] where Model='";
            string sqlcomm = "select *from [Danfossdata].[dbo].[";
            sqlcomm = sqlcomm + tablename + "]";
         //   MessageBox.Show(sqlcomm);
            SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
            DataSet thisDataset = new DataSet();
            try
            {
                SqlDap.Fill(thisDataset, "Testdata");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
             dt = thisDataset.Tables["Testdata"];
         //   DataView dv = new DataView(thisDataset.Tables["Testdata"]);
            SqlData.ItemsSource = dt.DefaultView;
            try
            {
                con.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
           


        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            DataTableToExcel(dt, "D:\\" + tablename + ".xls");

        }

        public static bool DataTableToExcel(DataTable dt, string path)
        {
            bool result = false;
            IWorkbook workbook = null;
            FileStream fs = null;
            IRow row = null;
            ISheet sheet;
            ICell cell = null;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    workbook = new HSSFWorkbook();
                    //  sheet = new ISheet();
                    sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表  
                                                           //sheet = workbook.CreateSheet("Sheet1");//创建一个名称为Sheet0的表  
                    int rowCount = dt.Rows.Count;//行数  
                    int columnCount = dt.Columns.Count;//列数  
                                                       //  MessageBox.Show(rowCount.ToString());
                                                       //  MessageBox.Show(columnCount.ToString());
                                                       //设置列头  
                    row = sheet.CreateRow(0);//excel第一行设为列头  
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //设置每行每列的单元格,  
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        // MessageBox.Show(DateTime.Now.Month.ToString());
                        //   row = sheet[l].CreateRow(i);
                        // MessageBox.Show(i.ToString());
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据  
                                                     //   MessageBox.Show(dt.Rows[i][j].ToString());
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }

                    }
                    // fs = File.Create(path)
                    using (fs = File.Create(path))
                    {
                        workbook.Write(fs);//向打开的这个xls文件中写入数据  
                        result = true;
                    }
                    fs.Close();
                }
                return result;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return false;
            }


        }










    }
}
