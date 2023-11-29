using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace PriJobAide
{
    public partial class FrmContext : Form
    {
        public string gName;
        public string gId;
        public string gExamFolder;
        public string gFileName;
        
        public FrmContext(string id, string examfolder,string filename)
        {
            
            gId = id;
            gExamFolder = examfolder;
            gFileName = filename;
            
            InitializeComponent();
        }
       
        private void FrmContext_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text != "")
            {
                try
                {
                    if (File.Exists(gExamFolder + @"\" + gFileName + ".txt"))
                    {
                        if (File.Exists(gExamFolder + @"\" + gFileName + ".txt"))
                        {
                            List<string> oldvalues = File.ReadAllLines(gExamFolder + @"\" + gFileName + ".txt").ToList();
                            oldvalues[1] = textBox1.Text;
                            oldvalues[4] = "Aguarda impressao.";
                            File.WriteAllLines(gExamFolder + @"\" + gFileName + ".txt", oldvalues.ToArray());
                        }

                        this.Close();
                    }
                }
                catch { }
                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
