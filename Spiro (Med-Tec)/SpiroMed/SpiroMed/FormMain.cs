using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SpiroMed
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
            MyUser = new ClassUser(0, new DataTable());
            ChangeUser(MyUser);
        }
        private ClassUser MyUser;
        public void ChangeUser(ClassUser NewUser)
        {
            MyUser = NewUser;
            if (NewUser.GetID == 0)
            {
                //mostrar Form Login
                foreach (Control ctl in panelMain.Controls)
                {
                    if (ctl is Form)
                    {
                        ((Form)ctl).Close();
                    }
                }
                FormLogin FrmLogin = new FormLogin(this);
                FrmLogin.TopLevel = false;
                FrmLogin.Dock = DockStyle.Fill;
                panelMain.Controls.Add(FrmLogin);
                FrmLogin.Show();

            }
            else
            {
                foreach (Control ctl in panelMain.Controls)
                {
                    if (ctl is Form)
                    {
                        ((Form)ctl).Close();
                    }
                }
                FormView FrmLogin = new FormView(this, NewUser);
                FrmLogin.TopLevel = false;
                FrmLogin.Dock = DockStyle.Fill;
                panelMain.Controls.Add(FrmLogin);
                FrmLogin.Show();
            }
 
        }
    
    }
}
