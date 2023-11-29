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
    public partial class FormView : Form
    {
        public FormView(Form _Parentes, ClassUser _ThisUser)
        {
            InitializeComponent();
            Parentes = _Parentes;
            ThisUser = _ThisUser;
        }
        private Form Parentes;
        private ClassUser ThisUser;
        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (this.tableLayoutPanel1.RowStyles[0].Height == 0)
            {
                this.tableLayoutPanel1.RowStyles[0].SizeType = SizeType.Absolute;
                this.tableLayoutPanel1.RowStyles[0].Height = 100;
            }
            else
            {
                this.tableLayoutPanel1.RowStyles[0].SizeType = SizeType.Absolute;
                this.tableLayoutPanel1.RowStyles[0].Height = 0;
            }
        }

        private void pesquisarToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
