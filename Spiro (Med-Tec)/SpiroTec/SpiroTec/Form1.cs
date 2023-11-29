using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Data.SqlServerCe;
using System.Drawing.Drawing2D;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices;

namespace SpiroTec
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Meds = new DataTable();
        }
        DataTable Meds;
        public static string GSPath
        {
            get
            {
                // 32-bit
                return System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath) + @"\Data\Dll\" + "gsdll32.dll";
            }
        }
        private void pesquisarToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void adquirirExameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (bckExamGetter.IsBusy == false)
            {
                OpenFileDialog fbd = new OpenFileDialog();
                fbd.Filter = "Pdf (.pdf)|*.pdf";
                fbd.Multiselect = false;
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK)
                {
                    FormProgress frmProgress = new FormProgress();
                    frmProgress.Owner = this;
                    frmProgress.StartPosition = FormStartPosition.CenterParent;
                    object[] Arguments = new object[] { frmProgress, fbd.FileName };
                    bckExamGetter.RunWorkerAsync(Arguments);
                    frmProgress.ShowDialog();
                }
            }
        }

        private void bckExamGetter_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] Arguments = (object[])e.Argument;
            FormProgress frmProgress = (FormProgress)Arguments[0];
            string FileName = (string)Arguments[1];
            try
            {
                frmProgress.Invoke(new Action(() => frmProgress.SetText("A verificar ficheiro.")));
            }
            catch { }
            if (FileInUse(FileName) == false)
            { 
            //LerPDF
                byte[] BinaryPdfData = File.ReadAllBytes(FileName);
                try
                {
                    frmProgress.Invoke(new Action(() => frmProgress.SetText("A ler ficheiro.")));
                }
                catch { }
                List<string> Content = TryReadPDF(BinaryPdfData, 1, "");
                if (Content.Count > 0)
                {
                    try
                    {
                        frmProgress.Invoke(new Action(() => frmProgress.SetText("A verificar conteudo.")));
                    }
                    catch { }
                    //Procurar Ientificadores
                    string MSGBOX = "";
                    foreach (string s in Content)
                    {
                        MSGBOX = MSGBOX + s + Environment.NewLine;
                    }
                    //MessageBox.Show(MSGBOX);
                    int Identificador1 = Content.FindIndex(x => x.Contains("Impresso por winspiroExpress"));
                    if (Identificador1 > 0)
                    {
                        try
                        {
                            frmProgress.Invoke(new Action(() => frmProgress.SetText("Identificado.")));
                        }
                        catch { }
                        //File.WriteAllLines("C:\\Spiro\\Spiro.txt",Content);
                        string NMarcacao = "";
                        try{
                        frmProgress.Invoke(new Action(() => frmProgress.SetText("A verificar numero de marcação.")));
                        }catch{}
                        #region Nmarcacao
                        try
                        {
                            int NMarcacao1 = Content.FindIndex(x => x.Contains("Sexo")) + 1;
                            string Converted = System.Text.RegularExpressions.Regex.Replace(Content[NMarcacao1], @"[\D-]", string.Empty);
                            //MessageBox.Show(Converted);
                            if (string.IsNullOrEmpty(Converted.Trim()))
                            {
                                //ShowDialog insert Marcacao
                                string NovoNMarcacao = "";
                                if (InputBox("Sem numero de Marcação", "Marcação:", ref NovoNMarcacao) == DialogResult.OK)
                                {
                                    string Converted2 = System.Text.RegularExpressions.Regex.Replace(NovoNMarcacao, @"[\D-]", string.Empty);
                                    if (string.IsNullOrEmpty(Converted2.Trim()))
                                    {
                                        MessageBox.Show("Numero de Marcação inválido. Ficheiro não processado.");
                                    }
                                    else
                                    {
                                        NMarcacao = Converted2;
                                    }
                                }

                            }
                            else
                            {
                                NMarcacao = Converted;
                            }


                        }
                        catch { } 
                        #endregion
                        //NMarcacao Done
                       
                        if (string.IsNullOrEmpty(NMarcacao.Trim()) == false)
                        {
                            #region Data
                            //Data

                            DateTime ExamDate = DateTime.Now;
                            int ExamDate1 = Content.FindIndex(x => x.Contains("Dados do Paciente")) - 1;
                            if (ExamDate1 > 0)
                            {

                                string ExamDateStr = Content[ExamDate1].Replace("Data ", "").Trim();
                                try
                                {
                                    ExamDate = DateTime.ParseExact(ExamDateStr, "dd-MM-yyyy   HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                                }
                                catch { }

                            } 
                            #endregion
                            // Data Done
                            string Nome = "";
                            
                            
                            int Nome1 = Content.FindIndex(x => x.Contains("Primeiro Nome"));
                            string[] splitNome = Content[Nome1].Split(new string [] {"Primeiro Nome"},StringSplitOptions.None);
                            try
                            {
                                Nome = splitNome[1];
                            }
                            catch { }
                            DataTable ExistingExam = new DataTable();
                            using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
                            {
                                string SQLExist = "Select * from TblSpiro WHERE NMarcacao = '" + NMarcacao + "'";
                                SqlCeCommand MyCmdExist = new SqlCeCommand(SQLExist, conn);
                                SqlCeDataAdapter daExist = new SqlCeDataAdapter(MyCmdExist);
                                try 
                                {
                                    daExist.Fill(ExistingExam);
                                }
                                catch { }
                            
                            }
                            
                            bool? Tipo = null; //Basal
                            if (ExistingExam.Rows.Count > 0)
                            #region Existe
                            {
                                // Perguntar se é Basal ou Pos BD
                                if (InputBoxPBD("Exame com este numero de Marcação já existe.", "O PDF é Pós-BD:", ref Tipo) == DialogResult.OK)
                                {

                                }

                                if (Tipo == null)
                                {
                                    MessageBox.Show("É necessário definir o tipo de exame a importar (Basal ou pós - BD)");
                                }
                                else
                                {
                                    string Rootfolder = System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath) + ExistingExam.Rows[0]["RootFolder"].ToString();
                                    string DBRootFolder = ExistingExam.Rows[0]["RootFolder"].ToString();
                                    string Prefix = "";
                                    if (Tipo.HasValue)
                                    {
                                        if (Tipo.Value == true)
                                        {
                                            Prefix = "POS_BD_";
                                        }
                                        if (Tipo.Value == false)
                                        {
                                            Prefix = "BASAL_";
                                        }
                                    }
                                    if (string.IsNullOrEmpty(Prefix.Trim()) == false)
                                    {
                                        byte[] AnonimPDFBinary = Anomizado(BinaryPdfData, ""," ID: " + NMarcacao);
                                        //ANONIMIZAR
                                        //COPIAR PARA ROOT
                                        System.IO.File.WriteAllBytes(Rootfolder + @"\" + Prefix + DateTime.Now.ToString("yyyyMMddHHmmss") + ".PDF", AnonimPDFBinary);
                                        //UPDATE
                                        using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
                                        {
                                            string SQLUpdate = "UPDATE TblSpiro SET Tipo = @Tipo, Sent = @Sent, Validated = @Validated WHERE ID_TblSpiro = " + Convert.ToInt32(ExistingExam.Rows[0]["ID_TblSpiro"]);
                                            SqlCeCommand MyCmdUpdate = new SqlCeCommand(SQLUpdate, conn);
                                            MyCmdUpdate.Parameters.AddWithValue("@Tipo", Tipo);
                                            MyCmdUpdate.Parameters.AddWithValue("@Sent", false);
                                            MyCmdUpdate.Parameters.AddWithValue("@Validated", false);
                                            try
                                            {
                                                if (conn.State == ConnectionState.Closed)
                                                { conn.Open(); }
                                                MyCmdUpdate.ExecuteNonQuery();
                                            }
                                            catch { }
                                        }
                                    }
                                    //Basal - Verificar existencia de PDF no remote root - BASAL_
                                    //Pos BD - Verificar existencia de PDF no remote Root - POS_BD_
                                    //se PDF encontrado perguntar se quer substituir
                                }
                            } 
                            #endregion
                            else
                            { 
                            // Create Remote 
                                Tipo = false;
                                string Prefix = "";
                                    if (Tipo.HasValue)
                                    {
                                        if (Tipo.Value == true)
                                        {
                                            Prefix = "POS_BD_";
                                        }
                                        if (Tipo.Value == false)
                                        {
                                            Prefix = "BASAL_";
                                        }
                                    }
                                string Rootfolder = UniqueDir(0,DateTime.Now);
                                string TempPath = System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath);
                                string DBRootFolder = Rootfolder.Replace(TempPath ,"");
                                bool DirCreated = true;

                                try { Directory.CreateDirectory(Rootfolder); }
                                catch { DirCreated = false; }
                                if (DirCreated == true)
                                {
                                    // ANOM
                                    byte[] AnomPDFBinary = Anomizado(BinaryPdfData, "", "ID: " + NMarcacao);
                                    System.IO.File.WriteAllBytes(Rootfolder + @"\" + Prefix + DateTime.Now.ToString("yyyyMMddHHmmss") + ".PDF", AnomPDFBinary);
                                    //INSERT
                                    using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
                                    {
                                        string SQLInsert = "INSERT INTO TblSpiro (Nome, Data, Tipo, NMarcacao, RootFolder, IDMed, Sent, Validated) VALUES (@Nome, @Data, @Tipo, @NMarcacao, @RootFolder, @IDMed, @Sent, @Validated)";
                                        SqlCeCommand MyCmdExist = new SqlCeCommand(SQLInsert, conn);
                                        MyCmdExist.Parameters.AddWithValue("@Nome", Nome);
                                        MyCmdExist.Parameters.AddWithValue("@Data", ExamDate);
                                        MyCmdExist.Parameters.AddWithValue("@Tipo", Tipo);
                                        MyCmdExist.Parameters.AddWithValue("@NMarcacao", NMarcacao);
                                        MyCmdExist.Parameters.AddWithValue("@RootFolder", DBRootFolder);
                                        MyCmdExist.Parameters.AddWithValue("@IDMed", 0);
                                        MyCmdExist.Parameters.AddWithValue("@Sent", false);
                                        MyCmdExist.Parameters.AddWithValue("@Validated", false);
                                        try
                                        {
                                            if (conn.State == ConnectionState.Closed)
                                            { conn.Open(); }
                                            MyCmdExist.ExecuteNonQuery();
                                        }
                                        catch { }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Erro na criação de diractoria Temp");
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Numero de marcação inválido. O ficheiro não irá ser processado.");
                        }

                    }
                    else
                    {
                        MessageBox.Show("O ficheiro selecionado não é uma espirometria válida.");
                    }
                }
                else
                {
                    MessageBox.Show("Não foi possível ler o PDF.");
                }
            }
            /*
             *  PdfContentByte cb = writer.DirectContent;
                cb.BeginText();
                BaseFont f_cn = BaseFont.CreateFont("c:\\windows\\fonts\\calibri.ttf", BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(f_cn, 6);
                cb.SetTextMatrix(475, 15);  //(xPos, yPos)
                cb.ShowText("Some text here and the Date: " + DateTime.Now.ToShortDateString());
                cb.EndText();
             * */
            //Verificar PDF
            //Extrair Valores
            //Valores Nulos - erro
            //ELSE
            //Checar Valores - Verificar se é Template de SPiro
            //Procurar NMarcacao
            //NMARCACAO Não ENCONTRADO - Form NMARCACAO
            //DATA NÃO ENCONTRADA - FORM DATA (ASSUMIR HOJE)
            //NMARCACAO EXISTE - FORM PERGUNTAR SE É POS BD
            //TIPO JÀ EXISTE PERGUNTAR SE QUER SUBSTITUIR POS BD OU BASAL
            //MARCACAO NÃO EXISTE - ASSUMIR BASAL
            //FAZER DIR ROOT
            //ANONIMIZAR
            //COPIAR PARA ROOT
            //ADICIONAR A DB
            //NMARCACAO - NMARCACAO
            //DATA - DATA
            //TIPO BASAL / POSBD
            //VALIDATED - FALSE
            //IDMED = 0

            Thread.Sleep(500);
            e.Result = frmProgress;
            
        }

        private void bckExamGetter_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FormProgress frmProgress = (FormProgress)e.Result;
            frmProgress.Close(); 
        }
        public bool FileInUse(string path)
        {
            try
            {
                //Just opening the file as open/create
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                {
                    //If required we can check for read/write by using fs.CanRead or fs.CanWrite
                }

                return false;
            }
            catch
            {
                return true;
            }
        }
        public List<string> TryReadPDF(byte[] binaryPdfData, int MaxPages, string password)
        {
            List<string> Contents = new List<string>();
            StringBuilder text = new StringBuilder();
            iTextSharp.text.pdf.PdfReader.unethicalreading = true;
            using (var pdfDataStream = new MemoryStream(binaryPdfData))
            {
                iTextSharp.text.pdf.PdfReader pdfreader;
                if (string.IsNullOrWhiteSpace(password))
                {
                    pdfreader = new iTextSharp.text.pdf.PdfReader(pdfDataStream);
                }
                else
                {
                    pdfreader = new iTextSharp.text.pdf.PdfReader(pdfDataStream, new System.Text.ASCIIEncoding().GetBytes(password));
                }
                using (pdfreader)
                {
                    if (pdfreader.NumberOfPages <= MaxPages)
                    {
                        MaxPages = pdfreader.NumberOfPages;
                    }
                    for (int page = 1; page <= MaxPages; page++)
                    {
                        //ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        string CurrentText = PdfTextExtractor.GetTextFromPage(pdfreader, page, new SimpleTextExtractionStrategy());
                        CurrentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(CurrentText))).Trim();
                        if (string.IsNullOrEmpty(CurrentText) == false)
                        {
                            text.Append(CurrentText + Environment.NewLine);
                        }
                    }
                }
                string[] lines = text.ToString().Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                foreach (string s in lines)
                {
                    if (string.IsNullOrWhiteSpace(s.Trim()) == false)
                    {
                        Contents.Add(s);
                    }
                }
            }
            return Contents;
        }
        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }
        private string DBConn()
        {
            return @"Data Source = " + System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath) + @"\DB.sdf";;
        }
        public static DialogResult InputBoxPBD(string title, string promptText, ref bool? value)
        {
            Form form = new Form();
            Label label = new Label();
            CheckBox textBox = new CheckBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.CheckState =  CheckState.Unchecked;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Checked;
            return dialogResult;
        }
        public byte[] Anomizado(byte[] binaryPdfData, string password, string msg)
        {
            byte[] newpdfBytes;
            iTextSharp.text.pdf.PdfReader.unethicalreading = true;
            iTextSharp.text.pdf.PdfReader pdfreader;
            using (var pdfDataStream = new MemoryStream(binaryPdfData))
            {
                if (string.IsNullOrWhiteSpace(password))
                {
                    pdfreader = new iTextSharp.text.pdf.PdfReader(pdfDataStream);
                }
                else
                {
                    byte[] newpassword = System.Text.ASCIIEncoding.ASCII.GetBytes(password);
                    pdfreader = new iTextSharp.text.pdf.PdfReader(pdfDataStream, newpassword);
                }

                using (MemoryStream ms = new MemoryStream())
                {
                    using (PdfStamper stamper = new PdfStamper(pdfreader, ms, '\0', true))
                    {

                        for (int page = 1; page <= pdfreader.NumberOfPages; page++)
                        {

                            int BitMapPositionX = 0;
                            int BitMapPositionY = 0;
                            int AbsolutePositionX = 0;
                            int AbsolutePositionY = 0;
                            
                                #region DefineBitmaps
                                
                                    try
                                    {
                                        BitMapPositionX = 160;//tamanho do Quadrado
                                        BitMapPositionY = 12;// tamanhao do quadrado
                                        AbsolutePositionX = 27;
                                        AbsolutePositionY = 688;

                                    }
                                    catch
                                    {
                                        BitMapPositionX = 0;
                                        BitMapPositionY = 0;
                                        AbsolutePositionX = 0;
                                        AbsolutePositionY = 0;

                                    }
                                
                                #endregion
                            /*
                                    PdfContentByte cb = PdfWriter. .DirectContent;
                cb.BeginText();
                BaseFont f_cn = BaseFont.CreateFont("c:\\windows\\fonts\\calibri.ttf", BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(f_cn, 6);
                cb.SetTextMatrix(475, 15);  //(xPos, yPos)
                cb.ShowText("Some text here and the Date: " + DateTime.Now.ToShortDateString());
                cb.EndText();
                             * */
             
                                #region StamImage
                                if (BitMapPositionX == 0 && BitMapPositionY == 0 && AbsolutePositionX == 0 && AbsolutePositionY == 0)
                                { }
                                else
                                {
                                    Bitmap bmp = new Bitmap(BitMapPositionX, BitMapPositionY);
                                    if (msg != "")
                                    {
                                        //RectangleF rectf = new RectangleF(70, 90, 90, 50);
                                        PointF firstLocation = new PointF(3f, 0f);
                                        Graphics g = Graphics.FromImage(bmp);
                                        g.Clear(System.Drawing.Color.White);
                                        g.SmoothingMode = SmoothingMode.HighQuality;
                                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                                        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                                        //g.DrawString(msg, new Font("Tahoma", 7), Brushes.Black, firstLocation,StringFormat.GenericDefault);

                                        g.Flush();

                                        //System.Drawing.Image NewBmp.Image = bmp;
                                    }
                                    System.Drawing.Image NewBmp = bmp;
                                    iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(NewBmp, System.Drawing.Imaging.ImageFormat.Bmp);
                                    //Sets the position that the image needs to be placed (ie the location of the text to be removed)
                                    image.SetAbsolutePosition(AbsolutePositionX, AbsolutePositionY);
                                    //Adds the image to the output pdf
                                    stamper.GetOverContent(page).AddImage(image, true);
                                    var cb = stamper.GetOverContent(page);
                                    var baseFont = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                                    //cb.SetTextMatrix(AbsolutePositionX, AbsolutePositionY);
                                    //Draw some text
                                    

                                    cb.BeginText();
                                    cb.SetFontAndSize(baseFont, 10);
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, msg, AbsolutePositionX + 4, AbsolutePositionY + 3, 0);
                                    //cb.ShowText(msg);
                                    cb.EndText();
                                    

                                }
                            
                                #endregion

                        }
                        if (string.IsNullOrWhiteSpace(password) == false)
                        {
                            stamper.SetEncryption(null, null, iTextSharp.text.pdf.PdfWriter.ALLOW_PRINTING | iTextSharp.text.pdf.PdfWriter.ALLOW_MODIFY_CONTENTS, iTextSharp.text.pdf.PdfWriter.DO_NOT_ENCRYPT_METADATA);
                        }
                        stamper.Writer.CloseStream = false;
                        stamper.Close();
                    }
                    

                    newpdfBytes = ms.ToArray();
                }
            }
            if (newpdfBytes != null)
            {
                return newpdfBytes;
            }
            else
            {
                return binaryPdfData;
            }
        }

        public string UniqueDir(int IDLocal, DateTime CreationTime)
        {
            string TempPath = System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath);
            string Return = IDLocal.ToString() + ".0_" + CreationTime.ToString("ddMMyyyyHHmmss");
            string UniqueDir = TempPath + @"\" + IDLocal.ToString() + ".0_" + CreationTime.ToString("yyyyMMddHmmss");
            if (Directory.Exists(UniqueDir) == true)
            {
                int i = 0;
                while (Directory.Exists(UniqueDir) == true)
                {
                    UniqueDir = TempPath + @"\" + IDLocal.ToString() + "." + i.ToString() + "_" + CreationTime.ToString("yyyyMMddHmmss");
                    i++;
                }
                Return = UniqueDir;

            }
            else
            { Return = UniqueDir; }
            return Return;
        }
        public void RefreshDB()
        {
            DateTime Start = dateTimePicker1.Value;
            DateTime End = dateTimePicker2.Value;
            object[] Arguments = new object[] { Start, End };
            if (bckUpdate.IsBusy == false)
            {
                bckUpdate.RunWorkerAsync(Arguments);
            }
        }

        private void bckUpdate_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] Arguments = (object[])e.Argument;
            DateTime Start = (DateTime)Arguments[0];
            DateTime End = (DateTime)Arguments[1];
              DataTable DtUpdate = new DataTable();
            using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
            {
              
                string Update = "Select * from TblSpiro WHERE (Data BETWEEN @DateStart AND @DateEnd) ORDER BY Data DESC";
                SqlCeCommand MyCmdUpdate = new SqlCeCommand(Update, conn);
                SqlCeDataAdapter daUpdate = new SqlCeDataAdapter(MyCmdUpdate);
                MyCmdUpdate.Parameters.AddWithValue("@DateStart", Start);
                MyCmdUpdate.Parameters.AddWithValue("@DateEnd", End.AddDays(1).AddSeconds(-1));
                daUpdate.Fill(DtUpdate);

            }
            e.Result = DtUpdate;

        }

        private void bckUpdate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            RefreshDB();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
            {
                string Update = "Select * from TblMed";
                SqlCeCommand MyCmdUpdate = new SqlCeCommand(Update, conn);
                SqlCeDataAdapter daUpdate = new SqlCeDataAdapter(MyCmdUpdate);
                daUpdate.Fill(Meds);
            }

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows  = false;
            dataGridView1.AllowUserToResizeColumns = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.CellContentClick += ButtonClick;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.CellValueChanged += new DataGridViewCellEventHandler(dgCellValueChanged);
            dataGridView1.CurrentCellDirtyStateChanged += new EventHandler(dgCurrentCellDirtyStateChanged);
            dataGridView1.CellFormatting += new DataGridViewCellFormattingEventHandler(MyExamDatagridView_CellFormat);
            DataGridViewRow row = new DataGridViewRow();
            row.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.RowTemplate = row;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.WindowState = FormWindowState.Maximized;

            
        }

        private void bckUpdate_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            DataTable Update = (DataTable)e.Result;
            dataGridView1.Columns.Clear();
            if (Update.Rows.Count > 0)
            {
                dataGridView1.DataSource = Update;

                foreach (DataColumn Dc in Update.Columns)
                {
                    var col = new DataGridViewColumn();
                    col.Name = Dc.ColumnName;
                    DataGridViewCell cell = new DataGridViewTextBoxCell();
                    col.CellTemplate = cell;
                    col.DataPropertyName = Dc.ColumnName;
                    if (Dc.ColumnName == "ID_TblSpiro" || Dc.ColumnName == "Tipo" || Dc.ColumnName == "RootFolder" || Dc.ColumnName == "IDMed" || Dc.ColumnName == "Sent" || Dc.ColumnName == "Validated")
                    {
                        col.Visible = false;
                    }
                    if (Dc.ColumnName == "Nome")
                    {
                        col.FillWeight = 200;
                    }
                    if (Dc.ColumnName == "Data")
                    {
                        col.FillWeight = 150;
                    }
                    if (Dc.ColumnName == "NMarcacao")
                    {
                        col.FillWeight = 100;
                    }

                    dataGridView1.Columns.Add(col);
                }
                
                #region Med Buton
                DataTable MedTbl = Meds.Copy();
                DataRow Nomed = MedTbl.NewRow();
                Nomed["ID_TblMed"] = 0;
                Nomed["Utilizador"] = "NoUser";
                Nomed["Nome"] = "SM";
                MedTbl.Rows.Add(Nomed);
                var Med = new DataGridViewComboBoxColumn();
                Med.DataSource = MedTbl;
                Med.DisplayMember = "Nome";
                Med.ValueMember = "ID_TblMed";
                Med.Name = "Medico";
                Med.HeaderText = "Médico";
                Med.Width = 120;
                Med.DataPropertyName = ("IDMed");
                #endregion
                dataGridView1.Columns.Add(Med);
                #region Med Buton
               
                var TipoStr = new DataGridViewTextBoxColumn();
                TipoStr.Name = "Tipologia";
                TipoStr.HeaderText = "Tipologia";
                TipoStr.FillWeight = 100;
                #endregion
                dataGridView1.Columns.Add(TipoStr);
                dataGridView1.Columns["ID_TblSpiro"].Visible = false;
                #region ValPri Buton
                var ValPri = new DataGridViewButtonColumn();
                ValPri.Name = "ButtonColumnView";
                ValPri.HeaderText = "Ver";
                ValPri.Text = "Ver Exame";
                ValPri.Name = "btnVer";
                ValPri.UseColumnTextForButtonValue = true;
                #endregion
                dataGridView1.Columns.Add(ValPri);
                #region ValPri Buton
                var ValPrint = new DataGridViewButtonColumn();
                ValPrint.Name = "ButtonColumnView";
                ValPrint.HeaderText = "Imprimir";
                ValPrint.Text = "Imprimir";
                ValPrint.Name = "btnPri";
                ValPrint.UseColumnTextForButtonValue = true;
                #endregion
                dataGridView1.Columns.Add(ValPrint);
            }
        }
        void dgCurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            if (senderGrid.IsCurrentCellDirty)
            {
                // This fires the cell value changed handler below
                senderGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }
        void MyExamDatagridView_CellFormat(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridView Target = (DataGridView)sender;
            if (e.RowIndex > -1 && e.ColumnIndex == Target.Columns["Tipologia"].Index)
            {
                if (Target["Tipo", e.RowIndex].Value != null)
                {
                    bool Tipo = Convert.ToBoolean(Target["Tipo", e.RowIndex].Value);
                    if (Tipo == true)
                    {
                        e.Value = "Simples + BD";
                    }
                    if (Tipo == false)
                    {
                        e.Value = "Simples";
                    }
                }
            }
            
        }
        private void dgCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn && e.RowIndex >= 0)
            {
                // My combobox column is the second one so I hard coded a 1, flavor to taste
                DataGridViewComboBoxCell cb = (DataGridViewComboBoxCell)senderGrid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (cb.Value != null || Convert.ToInt32(cb.Value) != Convert.ToInt32(((DataGridView)sender).Rows[e.RowIndex].Cells["IDMed"].Value))
                {

                    int ID = Convert.ToInt32(((DataGridView)sender).Rows[e.RowIndex].Cells["ID_TblSpiro"].Value);
                    using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
                    {

                        string Update = "UPDATE TblSpiro SET IDMed = @IDMed WHERE (ID_TblSpiro = " + ID + ")";
                        SqlCeCommand MyCmdUpdate = new SqlCeCommand(Update, conn);
                       
                        MyCmdUpdate.Parameters.AddWithValue("@IDMed", cb.Value);
                        if (conn.State == ConnectionState.Closed)
                        { conn.Open(); }
                        MyCmdUpdate.ExecuteNonQuery();

                    }
                    senderGrid.Invalidate();
                }
            }
        }
        private void ButtonClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                e.RowIndex >= 0)
            {
                string Column_Name = ((DataGridView)sender).Columns[e.ColumnIndex].Name;
                int ID = Convert.ToInt32(((DataGridView)sender).Rows[e.RowIndex].Cells["ID_TblSpiro"].Value);
                string RotFolder = ((DataGridView)sender).Rows[e.RowIndex].Cells["RootFolder"].Value.ToString();
                Console.WriteLine(Column_Name);
                if (Column_Name == "btnVal")//validar
                {
                    
                }
                if (Column_Name == "btnVer")//imprimir
                {

                    ExamViewcs ExamView = new ExamViewcs(RotFolder, ID);
                    ExamView.ShowDialog();
                   
                   


                }
                if (Column_Name == "btnMID")//MudarID
                {
                   
                }
            }
        }

        private void bckEmailSend_DoWork(object sender, DoWorkEventArgs e)
        {
            DataTable ToSend = new DataTable();
            using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
            {
                string Update = "SELECT TblSpiro.* from  TblSpiro WHERE (Sent = @Sent) AND (Validated = @Validated) AND (IDMed <> @IDMed)";
                SqlCeCommand MyCmdUpdate = new SqlCeCommand(Update, conn);
                MyCmdUpdate.Parameters.AddWithValue("@Sent", false);
                MyCmdUpdate.Parameters.AddWithValue("@Validated", false);
                MyCmdUpdate.Parameters.AddWithValue("@IDMed", 0);
                SqlCeDataAdapter daUpdate = new SqlCeDataAdapter(MyCmdUpdate);
                daUpdate.Fill(ToSend);
                Console.Write("Exames para enviar:" + ToSend.Rows.Count);
            }
            if (ToSend.Rows.Count > 0)
            {
                foreach (DataRow drTosend in ToSend.Rows)
                {
                    DataTable OBS = new DataTable();
                    using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
                    {
                        string Update = "SELECT TblObs.* from  TblObs WHERE ID_TblSpiro = " + Convert.ToInt32(drTosend["ID_TblSpiro"]);
                        SqlCeCommand MyCmdUpdate = new SqlCeCommand(Update, conn);
                        SqlCeDataAdapter daUpdate = new SqlCeDataAdapter(MyCmdUpdate);
                        daUpdate.Fill(OBS);
                       
                    }
                    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                    service.Credentials = new WebCredentials("joao.luizi", "joao2012.");
                    // Set the URL.
                    service.Url = new Uri(@"https://webmail.tecnifar.pt/ews/exchange.asmx");


                    // Create an email message and provide it with connection 
                    // configuration information by using an ExchangeService object named service.
                    EmailMessage message = new EmailMessage(service);

                    // Set properties on the email message.
                    message.Subject = "Utilizador = " + Meds.Select("ID_TblMed = " + Convert.ToInt32(drTosend["IDMed"]))[0]["Utilizador"].ToString();
                    string body = "";
                    body = body + "<DBValuesStart>" + Environment.NewLine;
                    foreach (DataColumn Dc in ToSend.Columns)
                    {
                        body = body + Dc.ColumnName + "=" + drTosend[Dc.ColumnName].ToString() + Environment.NewLine;
                    }
                    body = body + "<DBValuesEnd>" + Environment.NewLine;
                    body = body + "<OBSStart>" + Environment.NewLine;
                    foreach (DataRow dr in OBS.Rows)
                    {
                        body = body + dr["Obs"].ToString() + Environment.NewLine;
                    }
                    body = body + "<OBSEnd>" + Environment.NewLine;
                    message.Body = new MessageBody(BodyType.Text, body); ;
                    message.IsRead = false;
                    string RootFolder = System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath) + drTosend["RootFolder"].ToString();
                    System.IO.DirectoryInfo info = new System.IO.DirectoryInfo(RootFolder);
                    System.IO.FileInfo BASALfile = info.GetFiles("BASAL_*").OrderBy(p => p.CreationTime).FirstOrDefault();
                    System.IO.FileInfo POSBDfile = info.GetFiles("POS_BD*").OrderBy(p => p.CreationTime).FirstOrDefault();
                    if (BASALfile != null)
                    {
                        message.Attachments.AddFileAttachment(BASALfile.FullName);
                    }
                    if (POSBDfile != null)
                    {
                        message.Attachments.AddFileAttachment(POSBDfile.FullName);
                    }
                    Folder rootfolder = Folder.Bind(service, WellKnownFolderName.Inbox);
                    rootfolder.Load();
                    foreach (Folder folder in rootfolder.FindFolders(new FolderView(30)))
                    {
                        if (folder.DisplayName == "SpiroNValid")
                        {
                            var fid = folder.Id;
                            message.Save(fid);
                            //message.Move(fid);
                            using (SqlCeConnection conn = new SqlCeConnection(DBConn()))
                            {
                                string Update = "UPDATE TblSpiro  SET Sent = @Sent WHERE ID_TblSpiro = " + Convert.ToInt32(drTosend["ID_TblSpiro"]);
                                SqlCeCommand MyCmdUpdate = new SqlCeCommand(Update, conn);
                                MyCmdUpdate.Parameters.AddWithValue("@Sent", true);
                                if (conn.State == ConnectionState.Closed)
                                { conn.Open(); }
                                MyCmdUpdate.ExecuteNonQuery();

                            }
                        }
                    }



                }
            }
        }

        private void enviarExamesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (bckEmailSend.IsBusy == false)
            {
                bckEmailSend.RunWorkerAsync();
            }
        }

    }
}
