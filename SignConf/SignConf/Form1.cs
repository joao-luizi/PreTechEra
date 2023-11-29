using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SignConf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Doc = "";
            Sig = "";
            Folder = "";
            msDoc = null;
            msSig = null;
            DocImg = new List<Image>();
            SigImg = null;
        }
        string Doc;
        string Sig;
        string Folder;
        byte[] msDoc;
        System.IO.MemoryStream msSig;
        List<Image> DocImg;
        List<string> AditionalFiles = new List<string>();
        Image SigImg;
        private void btDoc_Click(object sender, EventArgs e)
        {
            OpenFileDialog fbd = new OpenFileDialog();
            fbd.Filter = "Pdf (.pdf)|*.pdf|JPEG Files (*.jpeg)|*.jpeg|PNG Files (*.png)|*.png|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif";
            fbd.Multiselect = false;
            DialogResult result = fbd.ShowDialog();

            if (result == DialogResult.OK)
            {
                txtDoc.Text = fbd.FileName;
                Doc = fbd.FileName;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //carregar Docs
            bool DocFailed = false;
            bool SigFailer = false;
            if (Doc.Trim() == "")
            {
                DocFailed = true;
            }
            else
            {
                try 
                {
                    msDoc = System.IO.File.ReadAllBytes(Doc);
                }
                catch { DocFailed = true; }
            }
            if (Sig.Trim() == "")
            {
                SigFailer = true;
            }
            else
            {
                try
                {
                    msSig = new System.IO.MemoryStream(System.IO.File.ReadAllBytes(Sig));
                }
                catch { SigFailer = true; }
            }
            if (SigFailer == false && DocFailed == false)
            {
                System.IO.FileInfo DocFi = new System.IO.FileInfo(Doc);
                if (DocFi.Extension == ".pdf")
                {
                    Image NewSig = Image.FromStream(msSig);
                    int percentageRedux = Convert.ToInt32(nmSigRedux.Value);
                    int positionX = Convert.ToInt32(nmPosX.Value);
                    int positionY = Convert.ToInt32(nmPosY.Value);
                    PDFHelper PdfHelp = new PDFHelper();
                    byte[] NewPDFByte = System.IO.File.ReadAllBytes(Doc);
                    pictureBox1.Image = PdfHelp.TryAnomizar(NewPDFByte, 1, "", NewSig, positionX, positionY, percentageRedux, "", new List<string>());
                    
                }
                else
                {
                    Image NewSig = Image.FromStream(msSig);
                    int percentageRedux = Convert.ToInt32(nmSigRedux.Value);
                    int positionX = Convert.ToInt32(nmPosX.Value);
                    int positionY = Convert.ToInt32(nmPosY.Value);
                    Image NewBase = Image.FromFile(Doc);
                    
                    var target = new Bitmap(NewBase.Width, NewBase.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    target.SetResolution(NewBase.HorizontalResolution, NewBase.VerticalResolution);
                    var graphics = Graphics.FromImage(target);
                    graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceOver; // this is the default, but just to be clear
                    if (positionX + NewSig.Width > target.Width)
                    {
                        positionX = target.Width - NewSig.Width;
                    }
                    if (positionY  > target.Height)
                    {
                        positionY = target.Height;
                    }
                    graphics.DrawImage(NewBase, 0, 0);
                    graphics.DrawImage(NewSig, positionX, positionY);
                    pictureBox1.Image = target;
                    
                    if (Folder != "")
                    { 
                    
                    }
                    graphics.Dispose();
                    //target.Save("filename.png", System.Drawing.Imaging.ImageFormat.Png);
                }
            }

            
        }

        private void btSig_Click(object sender, EventArgs e)
        {
            OpenFileDialog fbd = new OpenFileDialog();
            fbd.Filter = "JPEG Files (*.jpeg)|*.jpeg|PNG Files (*.png)|*.png|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif";
            fbd.Multiselect = false;
            DialogResult result = fbd.ShowDialog();

            if (result == DialogResult.OK)
            {
                txtSig.Text = fbd.FileName;
                Sig = fbd.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void btFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    txtFolder.Text = fbd.SelectedPath;
                    Folder = fbd.SelectedPath;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            AditionalFiles.Clear();
            OpenFileDialog fbd = new OpenFileDialog();
            fbd.Filter = "Pdf (.pdf)|*.pdf|JPEG Files (*.jpeg)|*.jpeg|PNG Files (*.png)|*.png|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif";
            fbd.Multiselect = true;
            DialogResult result = fbd.ShowDialog();

            if (result == DialogResult.OK)
            {
                AditionalFiles.AddRange(fbd.FileNames);
               
            }
            button4.Text = AditionalFiles.Count.ToString() + " Ficheiros Selecionados";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Dar Saida
            bool DocFailed = false;
            bool SigFailer = false;
            if (Doc.Trim() == "")
            {
                DocFailed = true;
            }
            else
            {
                try
                {
                    msDoc = System.IO.File.ReadAllBytes(Doc);
                }
                catch { DocFailed = true; }
            }
            if (Sig.Trim() == "")
            {
                SigFailer = true;
            }
            else
            {
                try
                {
                    msSig = new System.IO.MemoryStream(System.IO.File.ReadAllBytes(Sig));
                }
                catch { SigFailer = true; }
            }
            if (SigFailer == false && DocFailed == false)
            {
                System.IO.FileInfo DocFi = new System.IO.FileInfo(Doc);
                if (DocFi.Extension == ".pdf")
                {
                    Image NewSig = Image.FromStream(msSig);
                    int percentageRedux = Convert.ToInt32(nmSigRedux.Value);
                    int positionX = Convert.ToInt32(nmPosX.Value);
                    int positionY = Convert.ToInt32(nmPosY.Value);
                    PDFHelper PdfHelp = new PDFHelper();
                    byte[] NewPDFByte = System.IO.File.ReadAllBytes(Doc);
                    pictureBox1.Image = PdfHelp.TryAnomizar(NewPDFByte, 1, "", NewSig, positionX, positionY, percentageRedux, Folder, AditionalFiles);
                    

                }
                else
                {
                    Image NewSig = Image.FromStream(msSig);
                    int percentageRedux = Convert.ToInt32(nmSigRedux.Value);
                    int positionX = Convert.ToInt32(nmPosX.Value);
                    int positionY = Convert.ToInt32(nmPosY.Value);
                    Image NewBase = Image.FromFile(Doc);
                    
                    var target = new Bitmap(NewBase.Width, NewBase.Height);
                    target.SetResolution(NewBase.HorizontalResolution, NewBase.VerticalResolution);
                    var graphics = Graphics.FromImage(target);
                    graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceOver; // this is the default, but just to be clear
                    if (positionX + NewSig.Width > target.Width)
                    {
                        positionX = target.Width - NewSig.Width;
                    }
                    if (positionY > target.Height)
                    {
                        positionY = target.Height;
                    }
                    graphics.DrawImage(NewBase, 0, 0);
                    graphics.DrawImage(NewSig, positionX, positionY);
                    pictureBox1.Image = target;
                    target.Save(@"C:\Users\Joao\Documents\SigTeste\teste.bmp", System.Drawing.Imaging.ImageFormat.Bmp);
                    if (Folder != "")
                    {
                        List<byte[]> MergeJob = new List<byte[]>();
                        byte[] Img;
                        using (var ms = new System.IO.MemoryStream())
                        {
                            target.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                            Img = ms.ToArray();
                        }
                        using (iTextSharp.text.Document doc = new iTextSharp.text.Document(new iTextSharp.text.Rectangle(Math.Min(target.Width, target.Height), Math.Max(target.Height, target.Width))))
                        {
                            
                            using (System.IO.MemoryStream ms2 = new System.IO.MemoryStream())
                            {
                                using (iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, ms2))
                                {
                                    doc.Open();
                                    try
                                    {

                                        iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(Img);
                                        jpg.SetAbsolutePosition(0, 0);
                                        //jpg.ScaleAbsolute(doc.PageSize.Width, doc.PageSize.Height);
                                        jpg.ScaleToFit(doc.PageSize.Width, doc.PageSize.Height);
                                        //jpg.ScaleAbsoluteWidth(doc.PageSize.Width);
                                        doc.Add(jpg);
                                        
                                    }
                                    catch { }
                                    finally { doc.Close(); MergeJob.Add(ms2.ToArray()); }
                                }
                            }
                        }
                        foreach (string AditionalFile in AditionalFiles)
                        {
                            if (System.IO.Path.GetExtension(AditionalFile).Contains("pdf"))
                            {
                                MergeJob.Add(System.IO.File.ReadAllBytes(AditionalFile));
                            }
                            else
                            {
                                Image NewImg = Bitmap.FromFile(AditionalFile);
                                using (iTextSharp.text.Document doc = new iTextSharp.text.Document(new iTextSharp.text.Rectangle(NewImg.Width, NewImg.Height)))
                                {
                                    using (System.IO.MemoryStream ms2 = new System.IO.MemoryStream())
                                    {
                                        using (iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, ms2))
                                        {
                                            doc.Open();
                                            try
                                            {

                                                iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(AditionalFile);
                                                //jpg.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                                                jpg.SetAbsolutePosition(0, 0);
                                                jpg.ScaleAbsoluteHeight(doc.PageSize.Height);
                                                jpg.ScaleAbsoluteWidth(doc.PageSize.Width);
                                                doc.Add(jpg);
                                            }
                                            catch { }
                                            finally { MergeJob.Add(ms2.ToArray()); doc.Close(); }
                                        }
                                    }
                                }
                            }
                        }
                        PDFHelper MyPdfHelp = new PDFHelper();
                        if (MergeJob.Count == 1)
                        {
                            System.IO.File.WriteAllBytes(Folder + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf", MergeJob[0]);
                        }
                        else
                        {
                            byte[] CompletePDF = MyPdfHelp.MergePDF(MergeJob);
                            System.IO.File.WriteAllBytes(Folder + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf", CompletePDF);
                        }
                    }
                    graphics.Dispose();
                    //target.Save("filename.png", System.Drawing.Imaging.ImageFormat.Png);
                }
            }

            
        }
    }
}
