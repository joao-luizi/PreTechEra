using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlServerCe;

namespace SpiroTec
{
    public partial class ExamViewcs : Form
    {
        public ExamViewcs(string _RootFolder, int _ID_TblSpiro)
        {
            InitializeComponent();
            RootFolder = _RootFolder;
            ID_TblSpiro = _ID_TblSpiro;
            BASALImageList = new List<KeyValuePair<string, Image>>();
            POSBDImageList = new List<KeyValuePair<string, Image>>();
            ObsStart = new DataTable();
        }
        private string RootFolder;
        private int ID_TblSpiro;
        private DataTable ObsStart;
        List<KeyValuePair<string, Image>> BASALImageList;
        List<KeyValuePair<string, Image>> POSBDImageList;
        private void ChangeBASALIMAGELIST(List<KeyValuePair<string, Image>> newBASALImageList)
        {
            BASALImageList = newBASALImageList;
            Console.WriteLine("CHECKING TABPAGEs!");
            if (tabControl1.TabPages.ContainsKey("Basal"))
            {
                Console.WriteLine("TABPAGE FOUND!");

                foreach (Control ctl in tabControl1.TabPages["Basal"].Controls)
                {

                    if (ctl is Panel)
                            {

                                if (ctl.Name == "pnView")
                                {

                                    foreach (Control ctl2 in ctl.Controls)
                                    {
                                        if (ctl2 is PictureBox)
                                        {
                                            ctl2.Dispose();
                                        }
                                    }
                                    int i = 0;
                                    foreach (KeyValuePair<string, Image> InImageList in newBASALImageList)
                                    {
                                        i++;
                                        PictureBox ToAdd = new PictureBox();
                                        ToAdd.Name = "pic" + i;
                                        ToAdd.Location = new Point(0, 0);
                                        ToAdd.Width = ctl.Width;
                                        int ImagWidth = InImageList.Value.Width;
                                        int ImagHeigth = InImageList.Value.Height;
                                        int percentageX = (ctl.Width * 100) / ImagWidth;
                                        ToAdd.Height = (ImagHeigth * percentageX) / 100;
                                        ToAdd.Width = (ImagWidth * percentageX) / 100;
                                        ToAdd.SizeMode = PictureBoxSizeMode.Zoom;
                                        ToAdd.Location = new Point(0, 0);
                                        ToAdd.Image = InImageList.Value;
                                        ctl.Controls.Add(ToAdd);

                                    }
                                }
                            }
                   
                }
            }
        }
        private void ChangePOSBDIMAGELIST(List<KeyValuePair<string, Image>> newPOSBDImageList)
        {
            POSBDImageList = newPOSBDImageList;
            Console.WriteLine("CHECKING TABPAGEs!");
            if (tabControl1.TabPages.ContainsKey("posBD"))
            {
                Console.WriteLine("TABPAGE FOUND!");

                foreach (Control ctl in tabControl1.TabPages["posBD"].Controls)
                {

                    if (ctl is Panel)
                    {

                        if (ctl.Name == "pnView")
                        {

                            foreach (Control ctl2 in ctl.Controls)
                            {
                                if (ctl2 is PictureBox)
                                {
                                    ctl2.Dispose();
                                }
                            }
                            int i = 0;
                            foreach (KeyValuePair<string, Image> InImageList in newPOSBDImageList)
                            {
                                i++;
                                PictureBox ToAdd = new PictureBox();
                                ToAdd.Name = "pic" + i;
                                ToAdd.Location = new Point(0, 0);
                                ToAdd.Width = ctl.Width;
                                int ImagWidth = InImageList.Value.Width;
                                int ImagHeigth = InImageList.Value.Height;
                                int percentageX = (ctl.Width * 100) / ImagWidth;
                                ToAdd.Height = (ImagHeigth * percentageX) / 100;
                                ToAdd.Width = (ImagWidth * percentageX) / 100;
                                ToAdd.SizeMode = PictureBoxSizeMode.Zoom;
                                ToAdd.Location = new Point(0, 0);
                                ToAdd.Image = InImageList.Value;
                                ctl.Controls.Add(ToAdd);

                            }
                        }
                    }

                }
            }
        }
        private string DBConn()
        {
            return @"Data Source = " + System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath) + @"\DB.sdf"; ;
        }
        public static string GSPath
        {
            get
            {
                // 32-bit
                return System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath) + @"\Data\Dll\" + "gsdll32.dll";
            }
        }
        private void ExamViewcs_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            if (bckPrepareForm.IsBusy == false)
            {
                object[] Arguments = new object[] { RootFolder , ID_TblSpiro };
                bckPrepareForm.RunWorkerAsync(Arguments);
            }
            using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
            {
                string Update = "SELECT * from  TblObs WHERE (ID_TblSpiro = " + ID_TblSpiro + ")";
                SqlCeCommand MyCmdUpdate = new SqlCeCommand(Update, conn);
                SqlCeDataAdapter daUpdate = new SqlCeDataAdapter(MyCmdUpdate);
                daUpdate.Fill(ObsStart);
                if (ObsStart.Rows.Count > 0)
                {
                    txtObs.Text = ObsStart.Rows[0]["Obs"].ToString();
                }
            }
        }

        private void bckPrepareForm_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] Arguments = (object[])e.Argument;
            string MyRoot = (string)Arguments[0];
            int ID_TblSpiro = (int)Arguments[1];
            DataTable EsteExame = new DataTable();
            using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
            {
                string Select = "SELECT * from TblSpiro WHERE ID_TblSpiro = " + ID_TblSpiro;
                SqlCeCommand MyCmdSelect = new SqlCeCommand(Select, conn);
                SqlCeDataAdapter daSelect = new SqlCeDataAdapter(MyCmdSelect);
                daSelect.Fill(EsteExame);
                //OBSERVAÇÕES
            }
            if (EsteExame.Rows.Count > 0)
            {
                bckPrepareForm.ReportProgress(0, Convert.ToBoolean(EsteExame.Rows[0]["Tipo"]));
                string RootFolder = System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath) + EsteExame.Rows[0]["RootFolder"].ToString();
                System.IO.DirectoryInfo info = new System.IO.DirectoryInfo(RootFolder);
                System.IO.FileInfo BASALfile = info.GetFiles("BASAL_*").OrderBy(p => p.CreationTime).FirstOrDefault();
                System.IO.FileInfo POSBDfile = info.GetFiles("POS_BD*").OrderBy(p => p.CreationTime).FirstOrDefault();

                if (BASALfile != null)
                {
                    Console.WriteLine("BASAL OK");
                    byte[] byteBASAL = System.IO.File.ReadAllBytes(BASALfile.FullName);
                    List<KeyValuePair<string, Image>> BASALImageList = GetPDFCompleteImages(byteBASAL, "BASAL");
                    Console.WriteLine("BASAL" + BASALImageList.Count);
                    bckPrepareForm.ReportProgress(1, BASALImageList);
                    //Retirar Imagens
                }
                else
                { Console.WriteLine("BASAL NOK"); }
                if (POSBDfile != null)
                {
                    byte[] bytePOSBD = System.IO.File.ReadAllBytes(POSBDfile.FullName);
                    List<KeyValuePair<string, Image>> POSBDImageList = GetPDFCompleteImages(bytePOSBD, "POSBD");
                    bckPrepareForm.ReportProgress(2, POSBDImageList);
                    //Retirar Imagens
                }

            }
        }

        private void bckPrepareForm_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 0)
            {
                bool Tipo = (bool)e.UserState;
                if (Tipo == false)//basal
                {
                    TabPage Toadd = new TabPage();
                    Toadd.Name = "Basal";
                    Toadd.Text = "Basal";
                    Panel pn = new Panel();
                    pn.Dock = DockStyle.Fill;
                    pn.Name = "pnView";
                    pn.AutoScroll = true;
                    this.tabControl1.TabPages.Add(Toadd);
                    this.tabControl1.TabPages["Basal"].Controls.Add(pn);
                }
                if (Tipo == true)//PosBD
                {
                    TabPage Toadd = new TabPage();
                    Toadd.Name = "Basal";
                    Toadd.Text = "Basal";
                    Panel pn = new Panel();
                    pn.Dock = DockStyle.Fill;
                    pn.Name = "pnView";
                    pn.AutoScroll = true;
                    TabPage Toadd2 = new TabPage();
                    Toadd2.Name = "posBD";
                    Toadd2.Text = "posBD";
                    Panel pn2 = new Panel();
                    pn2.Dock = DockStyle.Fill;
                    pn2.Name = "pnView";
                    pn2.AutoScroll = true;
                    
                    this.tabControl1.TabPages.Add(Toadd);
                    this.tabControl1.TabPages["Basal"].Controls.Add(pn);

                   
                    this.tabControl1.TabPages.Add(Toadd2);
                    this.tabControl1.TabPages["posBD"].Controls.Add(pn2);

                   
                }
            }
            if (e.ProgressPercentage == 1)
            {
                List<KeyValuePair<string, Image>> BASALImageList = (List<KeyValuePair<string, Image>>)e.UserState;
                Console.WriteLine("REACHED REPORT PROGRESS BASAL");
                ChangeBASALIMAGELIST(BASALImageList);

            }
                if (e.ProgressPercentage == 2)
                {
                    List<KeyValuePair<string, Image>> PODBDImageList = (List<KeyValuePair<string, Image>>)e.UserState;
                    //Console.WriteLine("REACHED REPORT PROGRESS BASAL");
                    ChangePOSBDIMAGELIST(PODBDImageList);

                }
            
        }
        private List<KeyValuePair<string, System.Drawing.Image>> GetPDFCompleteImages(byte[] PDFData, string Tipologia)
        {
            List<KeyValuePair<string, System.Drawing.Image>> ImageList = new List<KeyValuePair<string, System.Drawing.Image>>();

            Ghostscript.NET.GhostscriptVersionInfo vesion = new Ghostscript.NET.GhostscriptVersionInfo(new Version(0, 0, 0), GSPath, string.Empty, Ghostscript.NET.GhostscriptLicense.GPL);
            using (var pdfDataStream = new System.IO.MemoryStream(PDFData))
            {
                using (var rasterizer = new Ghostscript.NET.Rasterizer.GhostscriptRasterizer())
                {
                    List<BackgroundWorker> ImagesBuilder = new List<BackgroundWorker>();
                    var background1 = new BackgroundWorker();
                    var background2 = new BackgroundWorker();
                    var background3 = new BackgroundWorker();
                    var background4 = new BackgroundWorker();
                    var background5 = new BackgroundWorker();
                    ImagesBuilder.Add(background1);
                    ImagesBuilder.Add(background2);
                    ImagesBuilder.Add(background3);
                    ImagesBuilder.Add(background4);
                    ImagesBuilder.Add(background5);
                    rasterizer.Open(pdfDataStream, vesion, true);
                    int I = 1;
                    int Final = rasterizer.PageCount;

                    while (I <= Final)
                    {
                        List<BackgroundWorker> NonWorkingList = ImagesBuilder.Where(x => x.IsBusy == false).ToList();
                        foreach (BackgroundWorker bck in NonWorkingList)
                        {
                            bck.DoWork += (s, ex) =>
                            {
                                int ThisBckPage = (int)ex.Argument;
                                System.Drawing.Image newImage = rasterizer.GetPage(300, 300, ThisBckPage);
                                KeyValuePair<string, System.Drawing.Image> ToAdd = new KeyValuePair<string, Image>(Tipologia, newImage);
                                ImageList.Add(ToAdd);
                            };
                            bck.RunWorkerAsync(I);
                            I++;
                            if (I > Final)
                            {
                                break;
                            }
                        }
                    }
                    int WorkingBck = ImagesBuilder.Where(x => x.IsBusy == true).ToList().Count;
                    while (WorkingBck > 0)
                    {
                        WorkingBck = ImagesBuilder.Where(x => x.IsBusy == true).ToList().Count;
                        if (WorkingBck == 0)
                        {
                            break;
                        }
                    }
                }
            }
            Console.WriteLine("IAMGELIST:" + ImageList.Count.ToString());
            return ImageList;
        }

        private void ExamViewcs_ResizeEnd(object sender, EventArgs e)
        {
            
        }

        private void tabControl1_SizeChanged(object sender, EventArgs e)
        {
            if (tabControl1.TabPages.ContainsKey("Basal"))
            {
                int MaxWidth = tabControl1.Width - 5;
                foreach (Control ctl in tabControl1.TabPages["Basal"].Controls)
                {
                    if (ctl is Panel)
                            {

                                if (ctl.Name == "pnView")
                                {
                                    foreach (Control pic in ctl.Controls)
                                    {
                                        if (pic is PictureBox)
                                        {
                                            int i = 0;
                                           
                                                i++;

                                                //((PictureBox)pic).Width = MaxWidth;
                                                int ImagWidth = ((PictureBox)pic).Image.Width;
                                                int ImagHeigth = ((PictureBox)pic).Image.Height;
                                                int percentageX = (MaxWidth * 100) / ImagWidth;
                                                ((PictureBox)pic).Height = (ImagHeigth * percentageX) / 100;
                                                ((PictureBox)pic).Width = (ImagWidth * percentageX) / 100;
                                        }
                                    }
                                }
                            }

                }
            }
            if (tabControl1.TabPages.ContainsKey("posBD"))
            {
                int MaxWidth = tabControl1.Width - 5;
                foreach (Control ctl in tabControl1.TabPages["posBD"].Controls)
                {
                    if (ctl is Panel)
                    {

                        if (ctl.Name == "pnView")
                        {
                            foreach (Control pic in ctl.Controls)
                            {
                                if (pic is PictureBox)
                                {
                                    int i = 0;

                                    i++;

                                    //((PictureBox)pic).Width = ctl.Width;
                                    int ImagWidth = ((PictureBox)pic).Image.Width;
                                    int ImagHeigth = ((PictureBox)pic).Image.Height;
                                    int percentageX = (MaxWidth * 100) / ImagWidth;
                                    ((PictureBox)pic).Height = (ImagHeigth * percentageX) / 100;
                                    ((PictureBox)pic).Width = (ImagWidth * percentageX) / 100;
                                }
                            }
                        }
                    }

                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
            {
                DataTable Obs = new DataTable();
                string Update = "SELECT * from  TblObs WHERE (ID_TblSpiro = " + ID_TblSpiro + ")";
                SqlCeCommand MyCmdUpdate = new SqlCeCommand(Update, conn);
                SqlCeDataAdapter daUpdate = new SqlCeDataAdapter(MyCmdUpdate);
                daUpdate.Fill(Obs);
                if (Obs.Rows.Count == 0)
                {
                    string Insert = "INSERT INTO TblObs (ID_TblSpiro, Obs) Values (@ID_TblSpiro, @Obs)";
                    SqlCeCommand MyCmdInsert = new SqlCeCommand(Insert, conn);
                    MyCmdInsert.Parameters.AddWithValue("@ID_TblSpiro", ID_TblSpiro);
                    MyCmdInsert.Parameters.AddWithValue("@Obs", txtObs.Text);
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();
                    }
                    MyCmdInsert.ExecuteNonQuery();

                }
                else
                {
                    string Update2 = "UPDATE  TblObs SET Obs = @Obs WHERE ID_TblSpiro = " +ID_TblSpiro;
                    SqlCeCommand MyCmdUpdate2 = new SqlCeCommand(Update2, conn);
                    MyCmdUpdate2.Parameters.AddWithValue("@Obs", txtObs.Text);
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();
                    }
                    MyCmdUpdate2.ExecuteNonQuery();
                }

            }
        }
    }
}
