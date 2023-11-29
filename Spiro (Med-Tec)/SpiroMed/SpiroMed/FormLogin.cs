using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlServerCe;

namespace SpiroMed
{
    public partial class FormLogin : Form
    {
        public FormLogin(Form _Parentes)
        {
            InitializeComponent();
            Parentes = _Parentes;
           
        }
        private Form Parentes;
        private string DBConn()
        {
            return @"Data Source = " + System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath) + @"\DB.sdf"; ;
        }
        private void txtDefault_SetText(TextBox ctl, string msg)
        {
            ctl.Text = msg;
            ctl.ForeColor = Color.Gray;
        }
        private void txtDefault_Leave(TextBox ctl, string msg, bool Pass)
        {
            if (ctl.Text.Trim() == "")
            { txtDefault_SetText(ctl, msg); if (Pass == true) { ctl.UseSystemPasswordChar = false; } }
        }
        private void txtDefault_Enter(TextBox ctl, bool Pass)
        {
            if (ctl.ForeColor == Color.Black)
                return;
            ctl.Text = "";
            ctl.ForeColor = Color.Black;
            if (Pass == true) { ctl.UseSystemPasswordChar = true; }
        }

        private void FormLogin_Load(object sender, EventArgs e)
        {
            txtDefault_SetText(txtUser, "Utilizador");
            txtDefault_SetText(txtPass, "Password");
        }

        private void txtUser_Leave(object sender, EventArgs e)
        {
            txtDefault_Leave((TextBox)sender, "Utilizador", false);
        }

        private void txtUser_Enter(object sender, EventArgs e)
        {
            txtDefault_Enter((TextBox)sender, false);
        }

        private void txtPass_Leave(object sender, EventArgs e)
        {
            txtDefault_Leave((TextBox)sender, "PassWord", true);
        }

        private void txtPass_Enter(object sender, EventArgs e)
        {
            txtDefault_Enter((TextBox)sender, true);
        }

        private void Login()
        {
            string Utilizador = txtUser.Text.Trim();
            string PassWord = txtPass.Text.Trim();
            if (string.IsNullOrWhiteSpace(Utilizador) || string.IsNullOrWhiteSpace(PassWord) || (PassWord == "PassWord") || (Utilizador == "Utilizador"))
            { }
            else
            {
                DataTable Med = new DataTable();
                using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
                {

                    string Update = "Select * from TblMed WHERE Utilizador = '" + Utilizador + "'";
                    SqlCeCommand MyCmdUpdate = new SqlCeCommand(Update, conn);
                    SqlCeDataAdapter daUpdate = new SqlCeDataAdapter(MyCmdUpdate);
                    daUpdate.Fill(Med);

                }
                //DataRow[] User = new  DataRow[10];//MyValue.GIniValuesDataSet.Tables["TblUsers"].Select("Utilizador = '" + Utilizador + "'");
                if (Med.Rows.Count > 0)
                {
                   
                    DataRow[] UserAndPass = Med.Select("PassWord = '" + PassWord + "'");
                    if (UserAndPass.Length > 0)
                    {
                        //Build User
                        ClassUser NewUser = new ClassUser(Convert.ToInt32(UserAndPass[0]["ID_TblMed"]), UserAndPass.CopyToDataTable());
                        //GiveUser To Parente
                        ((FormMain)Parentes).ChangeUser(NewUser);
                    }
                    else
                    {
                        //PassWord Incorrecta
                        lblPassWord.Visible = true;
                        timerClear.Start();
                    }
                }
                else
                {
                    //Utilzador não existe
                    lblUser.Visible = true;
                    timerClear.Start();
                }


            }
        }
        private void btLogin_Click(object sender, EventArgs e)
        {
            Login();
        }

        private void timerClear_Tick(object sender, EventArgs e)
        {
            lblUser.Visible = false;
            lblPassWord.Visible = false;
            timerClear.Stop();
        }

        private void btLogin_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                Login();
            }
        }

        private void txtUser_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                Login();
            }
        }

        private void txtPass_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                Login();
            }
        }
    }
}
