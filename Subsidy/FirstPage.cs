using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Subsidy
{
    public partial class FirstPage : Form
    {
        Public_Classes.ForDB MyClass = new Public_Classes.ForDB();
        Public_Classes.ForModule MyMenu = new Public_Classes.ForModule();
        public FirstPage()  
        {
            InitializeComponent();
        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripStatusLabel4_Click(object sender, EventArgs e)
        {
            statusStrip1.Items[3].Text = Public_Classes.ForDB.Login_Name;//如何取得用户名？
        }

        private void 合同检索ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyMenu.Show_Form(sender.ToString().Trim());
        }

        private void FirstPage_Load(object sender, EventArgs e)
        {

        }

        private void 到期提醒ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyMenu.Show_Form(sender.ToString().Trim());
        }

        private void 重新登录ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyMenu.Show_Form(sender.ToString().Trim());
        }

        private void 修改密码ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyMenu.Show_Form(sender.ToString().Trim());
        }

        private void FirstPage_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Public_Classes.ForDB.Login_Name != "")
            {
                DateTime logout = DateTime.Now;
                Public_Classes.ForModule.logout_time=Convert.ToString(logout);

                MyMenu.writeLog();
            }
        }

        private void toolStripStatusLabel3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            MyMenu.Show_Form("基本信息");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MyMenu.Show_Form("重新登录");
        }
    }
}
