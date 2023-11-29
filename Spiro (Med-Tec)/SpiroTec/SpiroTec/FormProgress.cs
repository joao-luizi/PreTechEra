using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SpiroTec
{
    public partial class FormProgress : Form
    {
        public FormProgress()
        {
            InitializeComponent();
           
        }
 
        private void FormProgress_Load(object sender, EventArgs e)
        {
           
        }
        public void SetText(string Msg)
        {
            this.label1.Text = Msg;
        }
    }
}
