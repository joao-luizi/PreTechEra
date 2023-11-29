using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpiroMed
{
    public class ClassUser
    {

        public ClassUser(int ID, System.Data.DataTable TblMed)
        {
            if (ID == 0)
            {
                ID_TblMed = 0;
                Sig = new System.Drawing.Bitmap(16, 16);
                Utilizador = "";
                Nome = "";
                Password = "";
            }
            else
            {
                if (TblMed.Select("ID_TblMed =" + ID).Length > 0)
                {
                    ID_TblMed = ID;
                    Utilizador = TblMed.Select("ID_TblMed =" + ID)[0]["Utilizador"].ToString();
                    try { Sig = System.Drawing.Image.FromFile(System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath) + @"\Data\Sig\" + Utilizador + ".bmp"); }
                    catch { }
                    Nome = TblMed.Select("ID_TblMed =" + ID)[0]["Nome"].ToString();
                    Password = TblMed.Select("ID_TblMed =" + ID)[0]["Password"].ToString();
                }
            }
        }
        private int ID_TblMed;
        public int GetID
        {
            get { return ID_TblMed; }
        }
        private System.Drawing.Image Sig;
        private string Utilizador;
        private string Nome;
        private string Password;
    }
}
