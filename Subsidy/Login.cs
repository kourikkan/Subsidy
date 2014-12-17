using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Subsidy
{
    public partial class Login : Form
    {
        Public_Classes.ForDB MyClass = new Subsidy.Public_Classes.ForDB();
        public Login()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            {
                if (textName.Text != "" & textPass.Text != "")
                {
                    SqlDataReader temDR = MyClass.getcom("select * from dt_User where Name='" + textName.Text.Trim() + "' and Pass='" + textPass.Text.Trim() + "'");
                    bool ifcom = temDR.Read();
                    if (ifcom)
                    {
                        Public_Classes.ForDB.Login_Name = textName.Text.Trim();
                        Public_Classes.ForDB.Login_ID = temDR.GetString(0);//getString()获取指定列的字符串值
                        Public_Classes.ForDB.My_con.Close();
                        Public_Classes.ForDB.My_con.Dispose();

                        this.DialogResult = DialogResult.OK;
                        Public_Classes.ForModule.login_time = DateTime.Now;
                        //this.Close();
                    }
                    else
                    {
                        MessageBox.Show("用户名或密码错误！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        textName.Text = "";
                        textPass.Text = "";
                    }
                    MyClass.con_close();
                }
                else
                    MessageBox.Show("请将登录信息添写完整！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void B__Click(object sender, EventArgs e)
        {
            /*if ((int)(this.Tag) == 1)
            {
                Public_Classes.ForDB.Login_n = 3;//?
                Application.Exit();
            }
            else
                if ((int)(this.Tag) == 2)
                    this.Close();
        */
            Application.Exit();
        }

        private void F_Login_Load(object sender, EventArgs e)//这个方法的作用是？
        {
            try
            {
                MyClass.con_open();  //连接数据库
                MyClass.con_close();
                textName.Text = "";
                textPass.Text = "";

            }
            catch
            {
                MessageBox.Show("数据库连接失败。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Application.Exit();
            }
        }
        //下面三个方法实现按下回车键在几个输入框中切换的效果
        private void textName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\r')
                textPass.Focus();
        }

        private void textPass_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\r')
               B_Login.Focus();　　　　　　　　　　　　　　　
        }

        private void F_Login_Activated(object sender, EventArgs e)
        {
            textName.Focus();
        }

        private void Login_Load(object sender, EventArgs e)
        {

        }
        }
    }
