using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;


namespace Subsidy
{
    public partial class MainFile : Form
    {

        public MainFile()
        {
            InitializeComponent();
        }

        #region  当前窗体的所有共公变量
        Public_Classes.ForDB MyClass = new Public_Classes.ForDB();
        Public_Classes.ForModule MyMC = new Public_Classes.ForModule();
        public static DataSet MyDS_Grid;
        public static string tem_Field = "";//与数据库交互的字段名
        public static string tem_Value = "";//字段的值
        public static string tem_ID = "";
        public static int hold_n = 0;
        #endregion

        #region 页面控件填入内容的格式
        private void MainFile_Load(object sender, EventArgs e)
        {
            //用dataGridView1控件显示数据
            MyDS_Grid = MyClass.getDataSet(Public_Classes.ForDB.AllSql, "tb_Main");
            dataGridView1.DataSource = MyDS_Grid.Tables[0];
            dataGridView1.AutoGenerateColumns = true;  //是否自动创建列
            dataGridView1.Columns[0].Width = 60;
            dataGridView1.Columns[1].Width = 80;

            for (int i = 11; i < dataGridView1.ColumnCount; i++)  //隐藏dataGridView1控件中不需要的列字段
            {
                dataGridView1.Columns[i].Visible = false;
            }

            MyMC.MaskedTextBox_Format(S_4);  //指定MaskedTextBox控件的格式
            MyMC.MaskedTextBox_Format(S_5);

            //MyMC.CoPassData(S_6, "tb_Flight");  //
            //MyMC.CoPassData(S_7, "tb_Interval");  //
            MyMC.CoPassData(S_8, "tb_Status");  //
            MyMC.CoPassData(S_9, "tb_Type");  //          

            textBox1.Text = Convert.ToString(MyDS_Grid.Tables[0].Rows.Count);
            Public_Classes.ForDB.AllSql = "Select * from tb_Main";
        }
        #endregion

        #region 显示数据库内容
        /// <summary>
        /// 动态读取指定的记录行，并进行显示.
        /// </summary>
        /// <param name="DGrid">DataGridView控件</param>
        /// <returns>返回string对象</returns> 
        public string Grid_Inof(DataGridView DGrid)
        {
            //当DataGridView控件的记录>1时，将当前行中信息显示在相应的控件上
            if (DGrid.RowCount > 1)
            {
                S_0.Text = DGrid[0, DGrid.CurrentCell.RowIndex].Value.ToString();
                S_1.Text = DGrid[1, DGrid.CurrentCell.RowIndex].Value.ToString();
                S_2.Text = DGrid[2, DGrid.CurrentCell.RowIndex].Value.ToString();
                S_3.Text = DGrid[3, DGrid.CurrentCell.RowIndex].Value.ToString();
                S_4.Text = MyMC.Date_Format(Convert.ToString(DGrid[4, DGrid.CurrentCell.RowIndex].Value).Trim());
                S_5.Text = MyMC.Date_Format(Convert.ToString(DGrid[5, DGrid.CurrentCell.RowIndex].Value).Trim());
                S_6.Text = DGrid[6, DGrid.CurrentCell.RowIndex].Value.ToString();
                S_7.Text = DGrid[7, DGrid.CurrentCell.RowIndex].Value.ToString();
                S_8.Text = DGrid[8, DGrid.CurrentCell.RowIndex].Value.ToString();
                S_9.Text = DGrid[9, DGrid.CurrentCell.RowIndex].Value.ToString();
                S_10.Text = DGrid[10, DGrid.CurrentCell.RowIndex].Value.ToString();
                //S_12.Text = DGrid[11, DGrid.CurrentCell.RowIndex].Value.ToString();

                return DGrid[1, DGrid.CurrentCell.RowIndex].Value.ToString();
            }
            else
            {
                //使用MyMeans公共类中的Clear_Control()方法清空指定控件集中的相应控件
                //MyMC.Clear_Control(TabControl1.TabPages[0].Controls);
                MyMC.Clear_Control(TabControl1.TabPages[1].Controls);
                tem_ID = "";
                return "";
            }
        }
        #endregion

        #region  按条件显示内容
        /// <summary>
        /// 通过公共变量动态进行查询.
        /// </summary>
        /// <param name="C_Value">条件值</param>
        public void Condition_Lookup(string C_Value)
        {
            MyDS_Grid = MyClass.getDataSet("Select * from tb_Main where " + tem_Field + "='" + tem_Value + "'", "tb_Main");
            dataGridView1.DataSource = MyDS_Grid.Tables[0];
            textBox1.Text = Convert.ToString(MyDS_Grid.Tables[0].Rows.Count);
        }
        #endregion

        #region 各种空函数
        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }
        private void S_7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void label14_Click(object sender, EventArgs e)
        {

        }
        private void label13_Click(object sender, EventArgs e)
        {

        }
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        #endregion

        #region 快速检索Combobox
        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                tem_Value = comboBox2.SelectedItem.ToString();
                Condition_Lookup(tem_Value);
            }
            catch
            {
                comboBox2.Text = "";
                MessageBox.Show("只能以选择方式查询。");
            }
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedIndex)  //向comboBox2控件中添加相应的查询条件
            {
                case 0:
                    {
                        comboBox2.Items.Clear();
                        comboBox2.Items.Add("保底补贴");
                        comboBox2.Items.Add("固定补贴");
                        tem_Field = "补贴类型";
                        break;

                    }
                case 1:
                    {
                        MyMC.CoPassData(comboBox2, "tb_Status");
                        tem_Field = "合同状态";
                        break;
                    }
            }
        }

        #endregion

        #region 各个页面初始设置
        private void tabControl1_Click(object sender, EventArgs e)
        {
            groupBox1.Enabled = true;
            Sut_Delete.Enabled = true;
            MyMC.Ena_Button(Sut_Add, Sut_Amend, Sut_Cancel, Sut_Save, 1, 1, 0, 0);
            if (TabControl1.SelectedTab.Name == "tabPage1")
            {
                hold_n = 0;  //恢复原始标识
                MyMC.Ena_Button(Sut_Add, Sut_Amend, Sut_Cancel, Sut_Save, 1, 1, 0, 0);  //
                groupBox1.Text = "";
                Sub_Table.Enabled = true;
            }
        }

        #endregion

        #region 浏览按钮
        private void N_First_Click_1(object sender, EventArgs e)
        {
            int ColInd = 0;
            if (dataGridView1.CurrentCell.ColumnIndex == -1 || dataGridView1.CurrentCell.ColumnIndex > 1)
                ColInd = 0;
            else
                ColInd = dataGridView1.CurrentCell.ColumnIndex;
            if ((((Button)sender).Name) == "N_First")
            {
                dataGridView1.CurrentCell = this.dataGridView1[ColInd, 0];
                MyMC.Ena_Button(N_First, N_Previous, N_Next, N_Cauda, 0, 0, 1, 1);
            }
            if ((((Button)sender).Name) == "N_Previous")
            {
                if (dataGridView1.CurrentCell.RowIndex == 0)
                {
                    MyMC.Ena_Button(N_First, N_Previous, N_Next, N_Cauda, 0, 0, 1, 1);
                }
                else
                {
                    dataGridView1.CurrentCell = this.dataGridView1[ColInd, dataGridView1.CurrentCell.RowIndex - 1];
                    MyMC.Ena_Button(N_First, N_Previous, N_Next, N_Cauda, 1, 1, 1, 1);
                }
            }
            if ((((Button)sender).Name) == "N_Next")
            {
                if (dataGridView1.CurrentCell.RowIndex == dataGridView1.RowCount - 2)//?
                {
                    MyMC.Ena_Button(N_First, N_Previous, N_Next, N_Cauda, 1, 1, 0, 0);
                }
                else
                {
                    dataGridView1.CurrentCell = this.dataGridView1[ColInd, dataGridView1.CurrentCell.RowIndex + 1];
                    MyMC.Ena_Button(N_First, N_Previous, N_Next, N_Cauda, 1, 1, 1, 1);
                }
            }
            if ((((Button)sender).Name) == "N_Cauda")
            {
                dataGridView1.CurrentCell = this.dataGridView1[ColInd, dataGridView1.RowCount - 2];
                MyMC.Ena_Button(N_First, N_Previous, N_Next, N_Cauda, 1, 1, 0, 0);
            }
        }
        private void N_Previous_Click_1(object sender, EventArgs e)
        {
            N_First_Click_1(sender, e);
        }

        private void N_Next_Click(object sender, EventArgs e)
        {
            N_First_Click_1(sender, e);
        }

        private void N_Cauda_Click(object sender, EventArgs e)
        {
            N_First_Click_1(sender, e);
        }

        #endregion

        #region 操作按钮
        private void Sut_Add_Click(object sender, EventArgs e)
        {
            MyMC.Clear_Control(TabControl1.TabPages[1].Controls);
            hold_n = 1;  //用于记录添加操作的标识
            MyMC.Ena_Button(Sut_Add, Sut_Amend, Sut_Cancel, Sut_Save, 0, 0, 1, 1);
            groupBox1.Text = "当前正在添加信息";
        }



        private void Sut_Amend_Click(object sender, EventArgs e)
        {
            hold_n = 2;  //用于记录修改操作的标识
            MyMC.Ena_Button(Sut_Add, Sut_Amend, Sut_Cancel, Sut_Save, 0, 0, 1, 1);
            groupBox1.Text = "当前正在修改信息";
        }

        private void Sut_Cancel_Click(object sender, EventArgs e)
        {
            hold_n = 0;  //恢复原始标识
            MyMC.Ena_Button(Sut_Add, Sut_Amend, Sut_Cancel, Sut_Save, 1, 1, 0, 0);
            groupBox1.Text = "";
            if (tem_Field == "")
                button1_Click(sender, e);
            else
                Condition_Lookup(tem_Value);
        }

        private void Sut_Save_Click(object sender, EventArgs e)
        {
            string All_Field = "ID,航线,补贴金额,合作方,起始日期,终止日期,机型,班期,合同状态,补贴类型,备注";
            try
            {
                if (hold_n == 1 || hold_n == 2) //判断当前是添加，还是修改操作
                {
                    Public_Classes.ForModule.ADDs = ""; //清空MyModule公共类中的ADDs变量
                    //用MyModule公共类中的Part_SaveClass()方法组合添加或修改的SQL语句
                    MyMC.Part_SaveClass(All_Field, S_0.Text.Trim(), TabControl1.TabPages[1].Controls, "S_", "tb_Main", 11, hold_n);
                    //如果ADDs变量不为空，则通过MyMeans公共类中的getsqlcom()方法执行添加、修改操作
                    if (Public_Classes.ForModule.ADDs != "")
                        MyClass.getsqlcom(Public_Classes.ForModule.ADDs);
                }
                Sut_Cancel_Click(sender, e);    //调用“取消”按钮的单击事件
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                Console.WriteLine(
               "\nStackTrace ---\n{0}", err.StackTrace);
            }
        }

        private void Sut_Delete_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount < 2) //判断dataGridView1控件中是否有记录
            {
                MessageBox.Show("数据表为空，不可以删除。");
                return;
            }
            //删除职工信息表中的当前记录
            MyClass.getsqlcom("Delete tb_Main where ID='" + S_0.Text.Trim() + "'");

            Sut_Cancel_Click(sender, e);    //调用“取消”按钮的单击事件
        }
        #endregion

        #region 其它
        private void button1_Click(object sender, EventArgs e)//显示全部数据
        {
            MyDS_Grid = MyClass.getDataSet(Public_Classes.ForDB.AllSql, "tb_Main");
            dataGridView1.DataSource = MyDS_Grid.Tables[0];
            dataGridView1.AutoGenerateColumns = true;  //是否自动创建列
            dataGridView1.Columns[0].Width = 60;
            dataGridView1.Columns[1].Width = 80;

            for (int i = 11; i < dataGridView1.ColumnCount; i++)  //隐藏dataGridView1控件中不需要的列字段
            {
                dataGridView1.Columns[i].Visible = false;
            }

            textBox1.Text = Convert.ToString(MyDS_Grid.Tables[0].Rows.Count);
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            MyMC.Show_DGrid(dataGridView1, TabControl1.TabPages[1].Controls, "S_");
        }//实现单击项目时控件内容随之改变

        private void Sub_Table_Click(object sender, EventArgs e)//导出到excel按钮 
        {
            MyMC.DataGridviewShowToExcel(dataGridView1, false);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MyMC.update1("Table_1");
            MyDS_Grid = MyClass.getDataSet("Select * from tb_Main where ID in (select ID from Table_1 where 剩余日期>0 AND 剩余日期<=60)", "tb_Main");
            dataGridView1.DataSource = MyDS_Grid.Tables[0];
            textBox1.Text = Convert.ToString(MyDS_Grid.Tables[0].Rows.Count);
        }//60天提醒按钮

        private void button4_Click(object sender, EventArgs e)
        {
            MyMC.update2("Table_1");
        }//更新Table_1三个字段
        private void button3_Click(object sender, EventArgs e)
        {
            MyMC.update1("Table_1");
            MyDS_Grid = MyClass.getDataSet("Select * from tb_Main where ID in (select ID from Table_1 where 月数之差>=0)", "tb_Main");
            dataGridView1.DataSource = MyDS_Grid.Tables[0];
            textBox1.Text = Convert.ToString(MyDS_Grid.Tables[0].Rows.Count);
        }//当月有效

        private void MainFile_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Public_Classes.ForDB.Login_Name != "")
            {
                DateTime logout = DateTime.Now;
                Public_Classes.ForModule.logout_time = Convert.ToString(logout);

                MyMC.writeLog();
            }
        }

        #endregion

        


    }
}
