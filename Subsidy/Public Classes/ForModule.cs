using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;

namespace Subsidy.Public_Classes
{
    class ForModule
    {
        #region  公共变量
        Public_Classes.ForDB ForDBClass = new Public_Classes.ForDB();   //声明ForDB类的一个对象，以调用其方法                
        public static string ADDs = "";  //用来存储添加或修改的SQL语句
        public static string FindValue = "";  //存储查询条件
        public static string User_ID = "";  //存储用户的ID编号
        public static string User_Name = "";    //存储用户名
        public static string logout_time="";
        public static DateTime login_time;
        #endregion

        #region  窗体的调用
        /// <summary>
        /// 窗体的调用.
        /// </summary>
        /// <param name="FrmName">调用窗体的Text属性值</param>
        public void Show_Form(string FrmName)
        {
            if (FrmName == "重新登录")  //判断当前要打开的窗体
            {
                Login Reload = new Subsidy.Login();
                Reload.Tag = 2;
                Reload.ShowDialog();    //显示窗体
                Reload.Dispose();
            }
            /*if (FrmName == "账号管理")
            {
                AccMana TheACC = new Subsidy.AccMana();
                TheACC.Text = "账号管理";
                TheACC.ShowDialog();
                TheACC.Dispose();
            }*/
             
            if (FrmName == "基本信息")
            {
                MainFile FileWork = new Subsidy.MainFile();
                FileWork.Text = "基本信息";
                FileWork.ShowDialog();
                FileWork.Dispose();
            }
            /*if (FrmName == "到期提醒")
            {
                DeadLine Alarm = new Subsidy.DeadLine();
                Alarm.Text = "到期提醒";   //设置窗体名称
                Alarm.Tag = 1; //设置窗体的Tag属性，用于在打开窗体时判断窗体的显示类形
                Alarm.ShowDialog();    //显示窗体
                Alarm.Dispose();
            }*/
        }
        #endregion

        #region  将日期转换成指定的格式
        /// <summary>
        /// 将日期转换成yyyy-mm-dd格式.
        /// </summary>
        /// <param name="NDate">日期</param>
        /// <returns>返回String对象</returns>
        public string Date_Format(string NDate)
        {
            string sm, sd;
            int y, m, d;
            try
            {
                y = Convert.ToDateTime(NDate).Year;
                m = Convert.ToDateTime(NDate).Month;
                d = Convert.ToDateTime(NDate).Day;
            }
            catch
            {
                return "";
            }
            if (m < 10)
                sm = "0" + Convert.ToString(m);
            else
                sm = Convert.ToString(m);
            if (d < 10)
                sd = "0" + Convert.ToString(d);
            else
                sd = Convert.ToString(d);
            return Convert.ToString(y) + "-" + sm + "-" + sd;
        }
        #endregion

        #region  遍历清空指定的控件
        /// <summary>
        /// 清空所有控件下的所有控件.
        /// </summary>
        /// <param name="Con">可视化控件</param>
        public void Clear_Control(Control.ControlCollection Con)
        {
            foreach (Control C in Con)
            { //遍历可视化组件中的所有控件
                if (C.GetType().Name == "TextBox")  //判断是否为TextBox控件
                    if (((TextBox)C).Visible == true)   //判断当前控件是否为显示状态
                        ((TextBox)C).Clear();   //清空当前控件
                if (C.GetType().Name == "MaskedTextBox")  //判断是否为MaskedTextBox控件
                    if (((MaskedTextBox)C).Visible == true)   //判断当前控件是否为显示状态
                        ((MaskedTextBox)C).Clear();   //清空当前控件
                if (C.GetType().Name == "ComboBox")  //判断是否为ComboBox控件
                    if (((ComboBox)C).Visible == true)   //判断当前控件是否为显示状态
                        ((ComboBox)C).Text = "";   //清空当前控件的Text属性值
                if (C.GetType().Name == "RichTextBox")  //判断是否为PictureBox控件
                    if (((RichTextBox)C).Visible == true)   //判断当前控件是否为显示状态
                        ((RichTextBox)C).Clear();   //清空当前控件的Image属性
            }
        }
        #endregion

        #region  设置MaskedTextBox控件的格式
        /// <summary>
        /// 将MaskedTextBox控件的格式设为yyyy-mm-dd格式.
        /// </summary>
        /// <param name="NDate">日期</param>
        /// <param name="ID">数据表名称</param>
        /// <returns>返回String对象</returns>
        public void MaskedTextBox_Format(MaskedTextBox MTBox)
        {
            MTBox.Mask = "0000-00-00";
            MTBox.ValidatingType = typeof(System.DateTime);
        }
        #endregion

        #region  向comboBox控件传递数据表中的数据
        /// <summary>
        /// 动态向comboBox控件的下拉列表添加数据.
        /// </summary>
        /// <param name="cobox">comboBox控件</param>
        /// <param name="TableName">数据表名称</param>
        public void CoPassData(ComboBox cobox, string TableName)
        {
            cobox.Items.Clear();
            SqlDataReader MyDR = ForDBClass.getcom("select * from " + TableName);
            if (MyDR.HasRows)
            {
                while (MyDR.Read())
                {
                    if (MyDR[1].ToString() != "" && MyDR[1].ToString() != null)
                        cobox.Items.Add(MyDR[1].ToString());
                }
            }
        }
        #endregion

        #region  用按钮控制数据记录移动时，改变按钮的可用状态
        /// <summary>
        /// 设置按钮是否可用.
        /// </summary>
        /// <param name="B1">首记录按钮</param>
        /// <param name="B2">上一条记录按钮</param>
        /// <param name="B3">下一条记录按钮</param>
        /// <param name="B4">尾记录按钮</param>
        /// <param name="NDate">B1标识</param>
        /// <param name="NDate">B2标识</param>
        /// <param name="NDate">B3标识</param>
        /// <param name="NDate">B4标识</param>
        public void Ena_Button(Button B1, Button B2, Button B3, Button B4, int n1, int n2, int n3, int n4)
        {
            B1.Enabled = Convert.ToBoolean(n1);
            B2.Enabled = Convert.ToBoolean(n2);
            B3.Enabled = Convert.ToBoolean(n3);
            B4.Enabled = Convert.ToBoolean(n4);
        }
        #endregion

        #region 保存添加或修改的信息.
        /// <summary>
        /// 保存添加或修改的信息.
        /// </summary>
        /// <param name="Sarr">数据表中的所有字段</param>
        /// <param name="ID1">第一个字段值</param>
        /// <param name="ID2">第二个字段值</param>
        /// <param name="Contr">指定控件的数据集</param>
        /// <param name="BoxName">要搜索的控件名称</param>
        /// <param name="TableName">数据表名称</param>
        /// <param name="n">控件的个数</param>
        /// <param name="m">标识，用于判断是添加还是修改</param>
        public void Part_SaveClass(string Sarr, string ID1, Control.ControlCollection Contr, string BoxName, string TableName, int n, int m)
        {
            string tem_Field = "", tem_Value = "";
            int p = 2;
            if (m == 1)
            {    //当m为1时，表示添加数据信息
                if (ID1 != "")
                { //根据参数值判断添加的字段
                    tem_Field = "ID";
                    tem_Value = "'" + ID1 + "'";
                    p = 1;
                }
                
            }
            else
                if (m == 2)
                {    //当m为2时，表示修改数据信息
                    if (ID1 != "")
                    { //根据参数值判断添加的字段
                        tem_Value = "ID='" + ID1 + "'";
                        p = 1;
                    }                 
                }

            if (m > 0)
            { //生成部份添加、修改语句
                string[] Parr = Sarr.Split(Convert.ToChar(','));
                for (int i = p; i < n; i++)
                {
                    string sID = BoxName + i.ToString();    //通过BoxName参数获取要进行操作的控件名称
                    foreach (Control C in Contr)
                    {   //遍历控件集中的相关控件
                        if (C.GetType().Name == "TextBox" | C.GetType().Name == "MaskedTextBox" | C.GetType().Name == "ComboBox" | C.GetType().Name == "RichTextBox")
                            if (C.Name == sID)
                            { //如果在控件集中找到相应的组件
                                string Ctext = C.Text;
                                if (C.GetType().Name == "MaskedTextBox")    //如果当前是MaskedTextBox控件
                                    Ctext = Date_Format(C.Text);    //对当前控件的值进行格式化
                                if (m == 1)
                                {    //组合SQL语句中insert的相关语句
                                    tem_Field = tem_Field + "," + Parr[i];
                                    if (Ctext == "")
                                        tem_Value = tem_Value + "," + "NULL";
                                    else
                                        tem_Value = tem_Value + "," + "'" + Ctext + "'";
                                }
                                if (m == 2)
                                {    //组合SQL语句中update的相关语句
                                    if (Ctext == "")
                                        tem_Value = tem_Value + "," + Parr[i] + "=NULL";
                                    else
                                        tem_Value = tem_Value + "," + Parr[i] + "='" + Ctext + "'";
                                }
                            }
                    }
                }
                ADDs = "";
                if (m == 1) //生成SQL的添加语句
                    ADDs = "insert into " + TableName + " (" + tem_Field + ") values(" + tem_Value + ")";
                if (m == 2) //生成SQL的修改语句
                    ADDs = "update " + TableName + " set " + tem_Value + " where ID='" + ID1 + "'";
            }
        }
        #endregion

        #region  将当前表的数据信息显示在指定的控件上
        /// <summary>
        /// 将DataGridView控件的当前记录显示在其它控件上.
        /// </summary>
        /// <param name="DGrid">DataGridView控件</param>
        /// <param name="GBox">GroupBox控件的数据集</param>
        /// <param name="TName">获取信息控件的部份名称</param>
        public void Show_DGrid(DataGridView DGrid, Control.ControlCollection GBox, string TName)
        {
            string sID = "";
            if (DGrid.RowCount > 0)
            {
                for (int i = 0; i < DGrid.ColumnCount; i++)
                {
                    sID = TName + i.ToString();
                    foreach (Control C in GBox)
                    {
                        if (C.GetType().Name == "RichTextBox" | C.GetType().Name == "TextBox" | C.GetType().Name == "MaskedTextBox" | C.GetType().Name == "ComboBox")
                            if (C.Name == sID)
                            {
                                if (C.GetType().Name != "MaskedTextBox")
                                    C.Text = DGrid[i, DGrid.CurrentCell.RowIndex].Value.ToString();
                                else
                                    C.Text = Date_Format(Convert.ToString(DGrid[i, DGrid.CurrentCell.RowIndex].Value).Trim());
                            }
                    }
                }
            }

        }
        #endregion
 
        #region DataGridView数据显示到Excel
        ///////////////////////////////////////////////////////////////////////////
        //                      DataGridView 导出到Excel                        ///
        ///////////////////////////////////////////////////////////////////////////
        public void GridToExcel(DataGridView gridView)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "导出Excel (*.xls)|*.xls|(*xlsx)|*.xlsx";
                saveFileDialog.FilterIndex = 1;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.CreatePrompt = true;
                saveFileDialog.Title = "导出Excel格式文档";
                saveFileDialog.ShowDialog();
                string strName = saveFileDialog.FileName;
                //创建工作薄
                if (strName.Length != 0)
                {
                    //toolStripProgressBar1.Visible = true;
                    //以下变量什么意思？
                    System.Reflection.Missing miss = System.Reflection.Missing.Value;
                    //下面这种定义可以用
                    //Microsoft.Office.Interop.Excel.ApplicationClass excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                    //下面这样定义，导出操作同样可以成功，但据说以下为接口，不太明白
                    Microsoft.Office.Interop.Excel.Application oexcel = new Microsoft.Office.Interop.Excel.Application();

                    oexcel.Application.Workbooks.Add(true);

                    oexcel.Visible = false;//若是true，则在导出的时候会显示EXcel界面

                    if (oexcel == null)
                    {
                        MessageBox.Show("EXCEL无法启动！(可能您没有安装EXCEL，或者版本与本程序不符)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    //下面三个语句是什么意思？
                    Microsoft.Office.Interop.Excel.Workbooks obooks = (Microsoft.Office.Interop.Excel.Workbooks)oexcel.Workbooks;
                    Microsoft.Office.Interop.Excel.Workbook obook = (Microsoft.Office.Interop.Excel.Workbook)(obooks.Add(miss));
                    Microsoft.Office.Interop.Excel.Worksheet osheet = (Microsoft.Office.Interop.Excel.Worksheet)obook.ActiveSheet;
                    osheet.Name = "数据";


                    //创建 Range ，方便释放资源（为何要释放资源？？）
                    Microsoft.Office.Interop.Excel.Range rans = (Microsoft.Office.Interop.Excel.Range)osheet.Cells;

                    //创建ran为了下面赋值时候使用
                    //Microsoft.Office.Interop.Excel.Range ran = null;

                    //添加表头
                    for (int i = 0; i < gridView.ColumnCount; i++)
                    {
                        //以下这句可用
                        //oexcel.Cells[1, i+1] = gridView.Columns[i].HeaderText.ToString();
                        //这是为单元格赋值的另一种方法
                        rans[1, i + 1] = gridView.Columns[i].HeaderText.ToString();
                    }

                    //填充数据
                    for (int i = 0; i < gridView.RowCount; i++)
                    {
                        //i为行，j为列
                        for (int j = 0; j < gridView.ColumnCount; j++)
                        {
                            ////注意：datagrid的引用方法是 datagrid1[列，行],即先列后行
                            if (gridView[j, i].Value.GetType() == typeof(string))
                            {
                                //oexcel.Cells[i + 2, j + 1] = "'" + gridView[j, i].Value.ToString();
                                rans[i + 2, j + 1] = "'" + gridView[j, i].Value.ToString();
                            }
                            else
                            {
                                //oexcel.Cells[i + 2, j + 1] = gridView[j, i].Value.ToString();
                                rans[i + 2, j + 1] = gridView[j, i].Value.ToString();
                            }
                        }
                        //toolStripProgressBar1.Value += 100 / gridView.RowCount;
                    }

                    //NAR(rans);  
                    osheet.SaveAs(strName, miss, miss, miss, miss, miss, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, miss, miss, miss);
                    //book.Close(false, miss, miss);
                    //books.Close();
                    //excel.Quit();
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(books);
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                    //下面销毁excel进程，为何无效呢？
                    obook.Close(false, miss, miss);
                    obooks.Close();
                    oexcel.Quit();
                   /* 
                    NAR(rans);
                    NAR(osheet);
                    NAR(obook);
                    NAR(obooks);
                    NAR(oexcel);
                    */
                    GC.Collect();
                    //GC.WaitForPendingFinalizers();        //不知作用是什么，有人说也须加上，可是经试不加也可以，不知为什么?
                    MessageBox.Show("数据已经成功导出!");
                    //toolStripProgressBar1.Value = 0;
                    System.Diagnostics.Process.Start(strName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 导出到EXCEL
        /// <summary>   
        /// 打开Excel并将DataGridView控件中数据导出到Excel  
        /// </summary>   
        /// <param name="dgv">DataGridView对象 </param>   
        /// <param name="isShowExcle">是否显示Excel界面 </param>   
        /// <remarks>  
        /// add com "Microsoft Excel 11.0 Object Library"  
        /// using Excel=Microsoft.Office.Interop.Excel;  
        /// </remarks>  
        /// <returns> </returns>   
        public bool DataGridviewShowToExcel(DataGridView dgv, bool isShowExcle)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "导出Excel (*.xls)|*.xls|(*xlsx)格式|*.xlsx";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = true;
            saveFileDialog.Title = "导出Excel格式文档";
            saveFileDialog.ShowDialog();
            string strName = saveFileDialog.FileName;

            if (dgv.Rows.Count == 0)
                return false;
            //建立Excel对象   
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            if (excel == null)
            {
                MessageBox.Show("Excel无法启动");
                return false;
            }
            Microsoft.Office.Interop.Excel.Workbook xlBook = excel.Workbooks.Add(true);
            Excel.Worksheet xSheet = (Excel.Worksheet)xlBook.ActiveSheet;
            //excel.SheetsInNewWorkbook = 1;
            //excel.Application.Workbooks.Add(true);
            excel.Visible = isShowExcle;

            //生成字段名称   
            for (int i = 0; i < dgv.ColumnCount-3; i++)
            {
                excel.Cells[1, i + 1] = dgv.Columns[i].HeaderText;
            }
            //填充数据   
            for (int i = 0; i < dgv.RowCount - 1; i++)
            {
                for (int j = 0; j < dgv.ColumnCount-3; j++)
                {
                    if (dgv[j, i].ValueType == typeof(string))
                    {
                        excel.Cells[i + 2, j + 1] = "'" + dgv[j, i].Value.ToString();
                    }
                    else
                    {
                        excel.Cells[i + 2, j + 1] = dgv[j, i].Value.ToString();
                    }
                }
            }
            //excel.Application.Workbooks.Add(true).Save();
            xlBook.SaveAs(strName,
       Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
       Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value,
       Missing.Value, Missing.Value); 
            xSheet = null;
            xlBook = null;
            excel.Quit(); //这一句是非常重要的，否则Excel对象不能从内存中退出 
            excel = null;
            return true;
        }
        #endregion

        #region 60天提醒
       public void update1(string TableName)
        {
            string tem_Value, tem_Value2, tem_Value3;

            tem_Value = "当前日期 = GETDATE()";
            tem_Value2 = "剩余日期=DATEDIFF(day,当前日期,endday)";
            tem_Value3 = "月数之差=DATEDIFF(MONTH,当前日期,endday)";


            ADDs = "update " + TableName + " set " + tem_Value;
            if (Public_Classes.ForModule.ADDs != "") //执行语句
                ForDBClass.getsqlcom(Public_Classes.ForModule.ADDs);

            ADDs = "update " + TableName + " set " + tem_Value2;
            if (Public_Classes.ForModule.ADDs != "") //执行语句
                ForDBClass.getsqlcom(Public_Classes.ForModule.ADDs);

            ADDs = "update " + TableName + " set " + tem_Value3;
            if (Public_Classes.ForModule.ADDs != "") //执行语句
                ForDBClass.getsqlcom(Public_Classes.ForModule.ADDs);

        }

       public void update2(string TableName)
       {
           string tem_Value;

           tem_Value = "endday = (select 终止日期 from tb_Main where Table_1.ID = tb_Main.ID)";          

           ADDs = "insert into Table_1(ID) select ID from tb_Main where ID not in (select ID from Table_1)";           
           try
           {
               if (Public_Classes.ForModule.ADDs != "") //执行语句
                   ForDBClass.getsqlcom(Public_Classes.ForModule.ADDs);
           }
           catch (Exception e)
           {
               MessageBox.Show(e.Message);
           }

           ADDs = "update " + TableName + " set " + tem_Value;
           if (Public_Classes.ForModule.ADDs != "") //执行语句
               ForDBClass.getsqlcom(Public_Classes.ForModule.ADDs);

       }

        
        
        /* public void setDate(string TableName)
        {
            string Ctext, tem_Value;

            DateTime now=DateTime.Now;
            Ctext = Date_Format(Convert.ToString(now).Trim());
            tem_Value = "当前日期='" + Ctext + "'";


            ADDs = "update " + TableName + " set " + tem_Value ;
            try
            {
                if (Public_Classes.ForModule.ADDs != "") //执行语句
                    ForDBClass.getsqlcom(Public_Classes.ForModule.ADDs);               
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            getDay(TableName);
        }

        public void getDay(string TableName)
        {
            string tem_Value;

            tem_Value = "剩余日期 = DATEDIFF(day,tb_Main.当前日期,tb_Main.终止日期)";

            ADDs = "update " + TableName + " set " + tem_Value + "where tb_Main.终止日期 is NOT NULL";

            try
            {
                if (Public_Classes.ForModule.ADDs != "") //执行语句
                    ForDBClass.getsqlcom(Public_Classes.ForModule.ADDs);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

        }

        public void getMonth(string TableName)
        {
            string tem_Value;

            tem_Value = "月数之差 = DATEDIFF(month,tb_Main.当前日期,tb_Main.终止日期)";

            ADDs = "update " + TableName + " set " + tem_Value + "where tb_Main.当前日期 is NOT NULL AND tb_Main.终止日期 is NOT NULL";

            try
            {
                if (Public_Classes.ForModule.ADDs != "") //执行语句
                    ForDBClass.getsqlcom(Public_Classes.ForModule.ADDs);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

        }*/
        #endregion

        #region 登陆日志
        public void writeLog()
        {
            StreamWriter sw = new StreamWriter("D:\\log.txt", true);
            sw.WriteLine(Public_Classes.ForDB.Login_Name + "\t" + login_time.ToString() + "\t" + logout_time);
            sw.Close();
        }
        #endregion
    }
}
        