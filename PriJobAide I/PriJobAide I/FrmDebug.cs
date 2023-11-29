using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Drawing.Printing;
using System.Xml;
using iTextSharp.text.pdf;
using Dotnet = System.Drawing.Image;
using PdfSharp;
using iTextSharp.text.pdf.parser;



namespace PriJobAide
{
    
    public partial class FrmDebug : Form
    {
        
 

        [DllImport("User32.dll")]
        private static extern int ShowWindow(int hwnd, int nCmdShow);
        [DllImport("user32.dll", EntryPoint = "FindWindowEx",CharSet = CharSet.Auto)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(String sClassName, String sAppName);
        [DllImport("user32.dll")]
        public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        static extern byte VkKeyScan(char ch);
        const int WM_KEYDOWN = 0x100;

        public Main internalMain;
        public FrmDebug(Main a)
        {
            internalMain = a;
            InitializeComponent();
        }

        private void FrmDebug_Load(object sender, EventArgs e)
        {
            
        }

        private void tmrupdate_Tick(object sender, EventArgs e)
        {
            try
            {
                this.textBox1.Text = internalMain.gPathLogs;
                this.textBox2a.Text = internalMain.gdebug.ToString();
                this.textBox3a.Text = internalMain.gstate.ToString();
                this.textBox4a.Text = internalMain.gDaystoFlush.ToString();
                this.textBox5a.Text = internalMain.gpathPDF;
                this.textBox6a.Text = internalMain.gpathSumatraPDF;
                this.textBox7a.Text = internalMain.gBookletPrinter;
                this.textBox8a.Text = internalMain.gBookletAudit;
                this.textBox9a.Text = internalMain.gBookletSecsToWait.ToString();
                this.textBox10a.Text = internalMain.gBookletProcess;
                this.textBox11a.Text = internalMain.gpathECG;

                this.textBox12a.Text = internalMain.gprintECGR.ToString();
                this.textBox13a.Text = internalMain.gECGRPrinter;
                this.textBox14a.Text = internalMain.gprintECGI.ToString();
                this.textBox15a.Text = internalMain.gECGIPrinter;
                this.textBox16a.Text = internalMain.gprintECGC.ToString();
                this.textBox17a.Text = internalMain.gECGCPrinter;
                this.textBox18a.Text = internalMain.gjoinECG.ToString();
                this.textBox19a.Text = internalMain.gpathECGTbl;

                this.textBox20a.Text = internalMain.gpathECO;
                this.textBox21a.Text = internalMain.gprintECOR.ToString();
                this.textBox22a.Text = internalMain.gECORPrinter;
                this.textBox23a.Text = internalMain.gprintECOI.ToString();
                this.textBox24a.Text = internalMain.gECOIPrinter;
                this.textBox25a.Text = internalMain.gprintECOC.ToString();
                this.textBox26a.Text = internalMain.gECOCPrinter;
                this.textBox27a.Text = internalMain.gjoinECO.ToString();
                this.textBox28a.Text = internalMain.gpathECOTbl;

                this.textBox29a.Text = internalMain.gpathHOLTER;
                this.textBox30a.Text = internalMain.gprintHOLTERR.ToString();
                this.textBox31a.Text = internalMain.gHOLTERRPrinter;
                this.textBox32a.Text = internalMain.gprintHOLTERI.ToString();
                this.textBox33a.Text = internalMain.gHOLTERIPrinter;
                this.textBox34a.Text = internalMain.gprintHOLTERC.ToString();
                this.textBox35a.Text = internalMain.gHOLTERCPrinter;
                this.textBox36a.Text = internalMain.gjoinHOLTER.ToString();
                this.textBox37a.Text = internalMain.gpathHOLTERTbl.ToString();
                this.textBox38a.Text = internalMain.gpathHOLTERTAdd;
                this.textBox39a.Text = internalMain.gpathHOLTERTPass;
                this.textBox40a.Text = internalMain.gHOLTERExternalAppPath;
                this.textBox41a.Text = internalMain.gHOLTERExternalApp;

                this.textBox42a.Text = internalMain.gpathMAPA;
                this.textBox43a.Text = internalMain.gprintMAPAR.ToString();
                this.textBox44a.Text = internalMain.gMAPARPrinter;
                this.textBox45a.Text = internalMain.gprintMAPAI.ToString();
                this.textBox46a.Text = internalMain.gMAPAIPrinter;
                this.textBox47a.Text = internalMain.gprintMAPAC.ToString();
                this.textBox48a.Text = internalMain.gMAPACPrinter;
                this.textBox49a.Text = internalMain.gjoinMAPA.ToString();
                this.textBox50a.Text = internalMain.gpathMAPATbl.ToString();

                this.textBox51a.Text = internalMain.gpathPE;
                this.textBox52a.Text = internalMain.gprintPER.ToString();
                this.textBox53a.Text = internalMain.gPERPrinter;
                this.textBox54a.Text = internalMain.gprintPEI.ToString();
                this.textBox55a.Text = internalMain.gPEIPrinter;
                this.textBox56a.Text = internalMain.gprintPEC.ToString();
                this.textBox57a.Text = internalMain.gPECPrinter;
                this.textBox58a.Text = internalMain.gjoinPE.ToString();
                this.textBox59a.Text = internalMain.gpathPETbl.ToString();
            }
            catch { }
        }
        public void starttimer()
        { this.tmrupdate.Start(); }
        private void button1_Click(object sender, EventArgs e)
        {
            int test = internalMain.gstate;
            if (test == 1)
            {
                internalMain.gstate = 0;
            }
            if (test == 0)
            {
                internalMain.gstate = 1;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bool test = internalMain.gdebug;
            if (test == true)
            {
                internalMain.gdebug = false;
            }
            if (test == false)
            {
                internalMain.gdebug = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Helper Myhelp = new Helper();
            List<string> PDFContent = new List<string>();
            List<string> EcgTbl = new List<string>();
            string pathECGtable = @"C:\PriJobAide\Bin\ECG.tbl";
            OpenFileDialog ofd = new OpenFileDialog();
            DialogResult dr = ofd.ShowDialog();
            
            if (dr == DialogResult.OK) // Test result.
            {
                string file = ofd.FileName;
                
                try
                {
                    FileInfo a = new FileInfo(file);
                    
                    
                    
                    PDFContent = Myhelp.ReadPdfFile(file);
                    EcgTbl = File.ReadAllLines(pathECGtable).ToList();

                    //ECGTBL: ler linha a linha e dividir entre ficheiros necessários para ser um R ou um I
                    //ECGTbl ler linha a linha e ler as strings em nome= e em ID= estas serão as strings a remover 
                    // o formato deverá ser nomeex="" (no casode sintra) nomepos=[1] (index do nome no caso de sintra)
                    // do PdfContent[1] deverá ser removido ""

                    // do mesmo modo o id deverá ter o Idex = "EVENTO:" (no caso de Sintra) e o IDpos=[index]
                    if (PDFContent.Count > 0)
                    {
                        string content = "";
                        foreach (string s in PDFContent)
                        {
                            content = content + Environment.NewLine + s;
                        }
                        MessageBox.Show(content);
                        File.WriteAllText(internalMain.gPathLogs + @"\" + "readpdf.txt", content);
                    }
                    
                }
                catch 
                {
                    
                }
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Helper Myhelp = new Helper();
            Myhelp.EmailCheck(internalMain.gpathHOLTERTAdd, internalMain.gpathHOLTERTPass, internalMain.gHOLTERSender, internalMain.gHOLTERReceiver, internalMain.gEmailInc);
        }

        private void button5_Click(object sender, EventArgs e)
        {


            Helper MyHelp = new Helper();
            IntPtr hwnd = IntPtr.Zero;
            IntPtr hwndchild = IntPtr.Zero;
            IntPtr hwndchild2 = IntPtr.Zero;
            string filename = @"J:\erro\R9359157.RPS";
            var p = new System.Diagnostics.Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            p.StartInfo.FileName = @"C:\Pyramis\Holter\HltRepViewer.exe";
            p.StartInfo.CreateNoWindow = false;
            p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
            p.StartInfo.Arguments = filename;
            p.StartInfo.RedirectStandardOutput = false;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.UseShellExecute = false;
            p.Start();
            
                bool firstwindow = true;
                
                
                while (p.MainWindowHandle == IntPtr.Zero)
                {
                    // Discard cached information about the process
                    // because MainWindowHandle might be cached.
                    if (FindWindow("#32770", "Vision Premier") != IntPtr.Zero)
                    {
                        hwndchild = FindWindow("#32770", "Vision Premier");
                        List<IntPtr> ChildWindows = MyHelp.GetAllChildrenWindowHandles(hwndchild, 10);
                        foreach (IntPtr a in ChildWindows)
                        {
                            string caption = MyHelp.GetWindowCaption(a);
                            if (caption == "Erro na abertura do relatório.")
                            { MessageBox.Show("erro"); }
                            else
                            {
                                if (caption.Contains("REGEDIT"))
                                { MessageBox.Show("ok"); }
                            }
                        }
                        PostMessage(hwndchild, 0x111, (IntPtr)0x00000002, (IntPtr)0x0010073E);
                        break;
                    }
                    p.Refresh();

                    Thread.Sleep(10);
                }
                int i = 0;

                hwnd = p.MainWindowHandle;
                while (p.MainWindowHandle != IntPtr.Zero)
                {
                   
                    if (firstwindow == true)
                    {
                        ShowWindow((int)hwnd, 2);
                        p.WaitForInputIdle();
                        firstwindow = false;
                        PostMessage(hwnd, 0x111, (IntPtr)0x0001835F, (IntPtr)0x00000000);
                    }
                    p.Refresh();
                    Thread.Sleep(10);
                    if (FindWindow("#32770", "Imprimir") != IntPtr.Zero)
                    {
                        hwndchild2 = FindWindow("#32770", "Imprimir");
                        ShowWindow((int)hwndchild, 2);
                        Thread.Sleep(1000);
                        PostMessage(hwndchild2, WM_KEYDOWN, (IntPtr)VkKeyScan((char)13), IntPtr.Zero);
                        Thread.Sleep(1000);
                        MyHelp.endapp(hwnd);
                        p.WaitForExit();
                    }
                    i++;
                    if (i == 5000)
                    {
                        MyHelp.endapp(hwnd);
                        p.WaitForExit();
                        
                    }
                }
                // Hide the window
                //ShowWindow((int)hwnd, 2);
                //p.WaitForInputIdle();
                //PostMessage(hwnd, 0x111, (IntPtr)0x0001835F, (IntPtr)0x00000000);
                
                //Thread.Sleep(500);
                //while (FindWindow("#32770", "Imprimir") == IntPtr.Zero)
                //{
                //    p.Refresh();
                //
                  //  Thread.Sleep(10);
                    //
                //}
                
                //hwndchild = FindWindow("#32770", "Imprimir");
                //ShowWindow((int)hwndchild, 2);
                //Thread.Sleep(1000);
                //PostMessage(hwndchild, WM_KEYDOWN, (IntPtr)VkKeyScan((char)13), IntPtr.Zero);
                //Thread.Sleep(1000);
                
                
                
                //MyHelp.endapp(hwnd);
                //p.WaitForExit();

                
            
            

                

            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string content = "";
            content = content + Environment.NewLine + internalMain.gdebug.ToString();
            content = content + Environment.NewLine + internalMain.gstate.ToString();
            content = content + Environment.NewLine + internalMain.gDaystoFlush.ToString();
            content = content + Environment.NewLine + internalMain.gpathPDF.ToString();
            content = content + Environment.NewLine + internalMain.gpathECG.ToString();
            content = content + Environment.NewLine + internalMain.gprintECGR.ToString();
            content = content + Environment.NewLine + internalMain.gECGRPrinter.ToString();
            content = content + Environment.NewLine + internalMain.gprintECGI.ToString();
            content = content + Environment.NewLine + internalMain.gECGIPrinter.ToString();
            content = content + Environment.NewLine + internalMain.gprintECGC.ToString();
            content = content + Environment.NewLine + internalMain.gECGCPrinter.ToString();
            content = content + Environment.NewLine + internalMain.gjoinECG.ToString();
            content = content + Environment.NewLine + internalMain.gpathECGTbl.ToString();
            content = content + Environment.NewLine + internalMain.gpathSumatraPDF.ToString();
            MessageBox.Show(content);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string printername = internalMain.gECGRPrinter;
            string sumatra = internalMain.gpathSumatraPDF;
            PrinterSettings printer = new PrinterSettings();
            printer.PrinterName = printername;
            OpenFileDialog ofd = new OpenFileDialog();
            DialogResult dr = ofd.ShowDialog();

            if (dr == DialogResult.OK) // Test result.
            {
                string file = ofd.FileName;

                if ((printer.IsValid == true) && File.Exists(sumatra))
                {
                    var p = new System.Diagnostics.Process();
                    ProcessStartInfo startInfo = new ProcessStartInfo();
                    p.StartInfo.FileName = sumatra;
                    //p.StartInfo.CreateNoWindow = false;
                    //p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                    p.StartInfo.Arguments = "-print-to "  + printername + " " + file;              
                    p.Start();
                    p.WaitForExit();
                    MessageBox.Show("-print-to " + printername + " " + file);
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //System.Xml.XmlTextReader reader = new System.Xml.XmlTextReader("c:\\teste\\Print-in.20130531.xml");
            //reader.XmlResolver = new XmlResolver("c:\\teste\\print-in-audit.xsl");
            Helper Myhelp = new Helper();
            XmlDocument xdcDocument = new XmlDocument();
            string path = internalMain.gBookletAudit;
            DirectoryInfo Audit = new DirectoryInfo(path);
            List<FileInfo> AuditList = Audit.GetFiles("*.xml").Where(x => x.LastWriteTime.Date == DateTime.Today.Date).ToList();
            List<FileInfo> OrderedAuditList = AuditList.OrderByDescending(x => x.LastWriteTime).ToList();
            if (OrderedAuditList.Count > 0)
            {
                xdcDocument.Load(OrderedAuditList[0].FullName);

                XmlElement xelRoot = xdcDocument.DocumentElement;
                string name = "";
                string time = "";
                string contents = "";
                foreach (XmlElement xndNode in xelRoot)
                {
                     name = xndNode.GetAttribute("DocumentName");
                     time = xndNode.GetAttribute("Time");

                    contents = name + " " + time;
                    //Your sql insert command will go here;
                }

                MessageBox.Show(contents);
                DateTime jobtime = Convert.ToDateTime(time);
                MessageBox.Show(time.ToString());
                MessageBox.Show((Myhelp.LastBookletPrintTime(path).AddSeconds(30) > DateTime.Now).ToString());
            }
            else { MessageBox.Show("Audit Empty"); }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string AuditIdentifier = "Print-in." + DateTime.Now.ToString("yyyyMMdd") + ".xml";
            MessageBox.Show(internalMain.gBookletAudit + @"\"+ AuditIdentifier);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Helper Myhelp = new Helper();
            List<string> FullTbl = new List<string>();
            List<string> PDFContent = new List<string>();
            List<string> ReportContent = new List<string>();
            List<string> ImgContent = new List<string>();
            
            OpenFileDialog ofd = new OpenFileDialog();
            DialogResult dr = ofd.ShowDialog();
            
            if (dr == DialogResult.OK) // Test result.
            {
                string file = ofd.FileName;

                try
                {
                   // List<string> dir = Directory.GetFiles(internalMain.gpathPDF).ToList();
                   // foreach (string s in dir)
                   // {
                        //file = s;
                        PDFContent = Myhelp.ReadPdfFile(file);
                        List<string> MyTables = new List<string>();
                        MyTables.Add(internalMain.gpathECGTbl);
                        MyTables.Add(internalMain.gpathECOTbl);
                        MyTables.Add(internalMain.gpathMAPATbl);
                        MyTables.Add(internalMain.gpathHOLTERTbl);
                        MyTables.Add(internalMain.gpathPETbl);
                        int ThisExamTyp = Myhelp.ExamType(PDFContent, MyTables);
                    
                        string ThisExamName = "default";
                        string ThisExamId = "default";
                        string ThisExameDate = "default";
                        File.WriteAllLines(@"C:\test.txt", PDFContent);
                        //MessageBox.Show(ThisExamTyp.ToString());
                        if (ThisExamTyp == 1 || ThisExamTyp == 2)
                        {
                            
                            string Temptbl = internalMain.gpathECGTbl;                     
                            FileInfo a = new FileInfo(file);
                            ThisExamName = Myhelp.GetName(ThisExamTyp, PDFContent, Temptbl);
                            ThisExamName = Myhelp.GetNormalizedName(ThisExamName);
                            ThisExamId = Myhelp.GetId(ThisExamTyp, PDFContent, Temptbl);
                            ThisExamId = Myhelp.GetNormalizedID(ThisExamId);
                            ThisExameDate = Myhelp.GetDate(ThisExamTyp, PDFContent, Temptbl);
                            //MessageBox.Show(ThisExamName + " " + ThisExamId);
                            string TargetPath = internalMain.gpathECG;
                            string type = "";
                            if (ThisExamTyp == 1)
                            { type = "R"; }
                            else { type = "I"; }
                            try
                            {
                                //if (Directory.Exists(TargetPath))
                                //{
                                //    a.MoveTo(TargetPath + @"\" + ThisExamName + "." + ThisExamId + "." + a.CreationTime.ToString("ddMMyyyHHmmss") + "." + type + "." + "0");
                                //}
                                MessageBox.Show(ThisExamName + Environment.NewLine + ThisExamId + Environment.NewLine + ThisExameDate + Environment.NewLine + type);
                            }
                            catch { }//log erro}
                        }
                        if (ThisExamTyp == 9 || ThisExamTyp == 10)
                        {
                            string Temptbl = internalMain.gpathPETbl;
                            ThisExamName = Myhelp.GetName(ThisExamTyp, PDFContent, Temptbl);
                            ThisExamName = Myhelp.GetNormalizedName(ThisExamName);
                            ThisExamId = Myhelp.GetId(ThisExamTyp, PDFContent, Temptbl);
                            ThisExamId = Myhelp.GetNormalizedID(ThisExamId);
                            ThisExameDate = Myhelp.GetDate(ThisExamTyp, PDFContent, Temptbl);
                            //MessageBox.Show(ThisExamName + " " + ThisExamId);
                            FileInfo a = new FileInfo(file);
                            string TargetPath = internalMain.gpathPE;
                            string type = "";
                            if (ThisExamTyp == 9)
                            { type = "R"; }
                            else { type = "I"; }
                            try
                            {
                                //if (Directory.Exists(TargetPath))
                                //{
                                //    a.MoveTo(TargetPath + @"\" + ThisExamName + "." + ThisExamId + "." + a.CreationTime.ToString("ddMMyyyHHmmss") + "." + type + "." + "0");
                                //}
                                MessageBox.Show(ThisExamName + Environment.NewLine + ThisExamId + Environment.NewLine + ThisExameDate + Environment.NewLine + type);
                            }
                            catch { }//log erro}
                        }
                        if (ThisExamTyp == 7 || ThisExamTyp == 8)
                        {
                            string Temptbl = internalMain.gpathMAPATbl;
                            ThisExamName = Myhelp.GetName(ThisExamTyp, PDFContent, Temptbl);
                            ThisExamName = Myhelp.GetNormalizedName(ThisExamName);
                            ThisExamId = Myhelp.GetId(ThisExamTyp, PDFContent, Temptbl);
                            ThisExamId = Myhelp.GetNormalizedID(ThisExamId);
                            ThisExameDate = Myhelp.GetDate(ThisExamTyp, PDFContent, Temptbl);
                            //MessageBox.Show(ThisExamName + " " + ThisExamId);
                            FileInfo a = new FileInfo(file);
                            string TargetPath = internalMain.gpathMAPA;
                            string type = "";
                            if (ThisExamTyp == 7)
                            { type = "R"; }
                            else { type = "I"; }
                            try
                            {
                                //if (Directory.Exists(TargetPath))
                                //{
                                //    a.MoveTo(TargetPath + @"\" + ThisExamName + "." + ThisExamId + "." + a.CreationTime.ToString("ddMMyyyHHmmss") + "." + type + "." + "0");
                                //}
                                MessageBox.Show(ThisExamName + Environment.NewLine + ThisExamId + Environment.NewLine + ThisExameDate + Environment.NewLine + type);
                            }
                            catch { }//log erro}
                        }
                        if (ThisExamTyp == 3 || ThisExamTyp == 4)
                        {
                            string Temptbl = internalMain.gpathECOTbl;
                            ThisExamName = Myhelp.GetName(ThisExamTyp, PDFContent, Temptbl);
                            ThisExamName = Myhelp.GetNormalizedName(ThisExamName);
                            ThisExamId = Myhelp.GetId(ThisExamTyp, PDFContent, Temptbl);
                            ThisExamId = Myhelp.GetNormalizedID(ThisExamId);
                            ThisExameDate = Myhelp.GetDate(ThisExamTyp, PDFContent, Temptbl);
                            //MessageBox.Show(ThisExamName + " " + ThisExamId);
                            FileInfo a = new FileInfo(file);
                            string TargetPath = internalMain.gpathECO;
                            string type = "";
                            if (ThisExamTyp == 3)
                            { type = "R"; }
                            else { type = "I"; }
                            try
                            {
                                //if (Directory.Exists(TargetPath))
                                //{
                                //    a.MoveTo(TargetPath + @"\" + ThisExamName + "." + ThisExamId + "." + a.CreationTime.ToString("ddMMyyyHHmmss") + "." + type + "." + "0");
                                //}
                                MessageBox.Show(ThisExamName + Environment.NewLine + ThisExamId + Environment.NewLine + ThisExameDate + Environment.NewLine + type);
                            }
                            catch { }//log erro}
                        }
                        if (ThisExamTyp == 5 || ThisExamTyp == 6)
                        {
                            string Temptbl = internalMain.gpathHOLTERTbl;
                            ThisExamName = Myhelp.GetName(ThisExamTyp, PDFContent, Temptbl);
                            ThisExamName = Myhelp.GetNormalizedName(ThisExamName);
                            ThisExamId = Myhelp.GetId(ThisExamTyp, PDFContent, Temptbl);
                            ThisExamId = Myhelp.GetNormalizedID(ThisExamId);
                            ThisExameDate = Myhelp.GetDate(ThisExamTyp, PDFContent, Temptbl);
                            //MessageBox.Show(ThisExamName + " " + ThisExamId);
                            FileInfo a = new FileInfo(file);
                            string TargetPath = internalMain.gpathHOLTER;
                            string type = "";
                            if (ThisExamTyp == 5)
                            { type = "R"; }
                            else { type = "I"; }
                            try
                            {
                                //if (Directory.Exists(TargetPath))
                                //{
                                //    a.MoveTo(TargetPath + @"\" + ThisExamName + "." + ThisExamId + "." + a.CreationTime.ToString("ddMMyyyHHmmss") + "." + type + "." + "0");
                                //}
                                MessageBox.Show(ThisExamName + Environment.NewLine + ThisExamId + Environment.NewLine + ThisExameDate + Environment.NewLine + type);
                            }
                            catch { }//log erro}
                        }

                    //}
                }
                catch { }


            }
        }
        
        private void button12_Click(object sender, EventArgs e)
        {
            string pathtoanalyze = internalMain.gpathECG;
            DataTable ECG = new DataTable();
            Helper Myhelp = new Helper();
            List<ListViewItem> oldlist = new List<ListViewItem>();
            List<ListViewItem> newlist = new List<ListViewItem>();
            foreach (ListViewItem lstitem in listView1.Items)
            {
                
                oldlist.Add(lstitem); 
            
            }
            
            
            

            DirectoryInfo di = new DirectoryInfo(pathtoanalyze);
            List<FileInfo> ECGFiles = di.GetFiles("*.pdf", SearchOption.TopDirectoryOnly).ToList();
            List<string> Anotherfile = Directory.GetFiles(pathtoanalyze).ToList();
            List<FileInfo> orderedList = ECGFiles.OrderByDescending(x => x.CreationTime).ToList();
            bool dontjoin = false;
            foreach (FileInfo a in orderedList)
            {

                bool b = Myhelp.isnormalized(a.Name);
                dontjoin = false;
                if (b == true)
                {
                    //MessageBox.Show(a.Name);
                    string[] split = a.Name.Split('.');
                    ListViewItem newCustomersRow = new ListViewItem(split[0]);
                    string[] subitems = new string[5];
                    subitems[0] = split[1];

                    //newCustomersRow.SubItems.Add(split[1]);
                    if (split[3] == "R")
                    {
                        bool imagexits = false;
                        subitems[1] = "sim";
                        foreach (string check in Anotherfile)
                        {

                            if (imagexits == false)
                            {
                                if (check.Contains(split[0]) && check.Contains(split[1]) && check.Contains(".I."))
                                {
                                    imagexits = true; subitems[2] = "sim";
                                }
                            }
                        }
                        if (imagexits == false)
                        {
                            subitems[2] = "não";
                        }
                    }
                    if (split[3] == "I")
                    {
                        bool imagexits = false;
                        subitems[2] = "sim";
                        foreach (string check in Anotherfile)
                        {

                            if (imagexits == false)
                            {
                                if (check.Contains(split[0]) && check.Contains(split[1]) && check.Contains(".R."))
                                {
                                    imagexits = true; subitems[1] = "sim"; dontjoin = true;
                                }
                            }
                        }
                        if (imagexits == false)
                        {
                            subitems[1] = "não";
                        }

                    }
                    subitems[3] = DateTime.ParseExact(split[2], "ddMMyyyyHHmmss", null).ToString("dd-MM-yyyy HH:mm:ss");
                    if (split[4] == "0")
                    { subitems[4] = "Aguarda impressão"; }

                    newCustomersRow.SubItems.AddRange(subitems);
                    if (dontjoin == false)
                    {
                        newlist.Add(newCustomersRow);
                    }
                }
                else { }
            }
            //compare old with new
            if (oldlist.Count > 0)
            {
                if (oldlist.Count == newlist.Count)
                {
                    bool needsupdate = false;
                    foreach (ListViewItem olditem in oldlist)
                    {
                        if (needsupdate == false)
                        {
                            foreach (ListViewItem newitem in newlist)
                            {
                                if (olditem.SubItems[0].Text + olditem.SubItems[1].Text == newitem.SubItems[0].Text + newitem.SubItems[1].Text)
                                {
                                    if (olditem.SubItems[2].Text + olditem.SubItems[3].Text == newitem.SubItems[2].Text + newitem.SubItems[3].Text)
                                    { }
                                    else
                                    { needsupdate = true; }
                                }
                            }
                        }
                    }
                    if (needsupdate == true)
                    { listView1.Items.Clear(); listView1.Items.AddRange(newlist.ToArray()); }
                }
                else
                { listView1.Items.Clear(); listView1.Items.AddRange(newlist.ToArray()); }
                
            }
            else
            { listView1.Items.AddRange(newlist.ToArray()); }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            //System.Xml.XmlTextReader reader = new System.Xml.XmlTextReader("c:\\teste\\Print-in.20130531.xml");
            //reader.XmlResolver = new XmlResolver("c:\\teste\\print-in-audit.xsl");
            DateTime timeofprint = DateTime.Now;
            
            string nametofind = "54516.pdf";
            XmlDocument xdcDocument = new XmlDocument();
            string path = @"C:\ProgramData\VS\";
            DirectoryInfo Audit = new DirectoryInfo(path);
            List<FileInfo> AuditList = Audit.GetFiles("*.xml").Where(x => x.LastWriteTime.Date == DateTime.Today.Date).ToList();
            List<FileInfo> OrderedAuditList = AuditList.OrderByDescending(x => x.LastWriteTime).ToList();
            if (OrderedAuditList.Count > 0)
            {
                xdcDocument.Load(OrderedAuditList[0].FullName);

                XmlElement xelRoot = xdcDocument.DocumentElement;


                string contents = "";
                foreach (XmlElement xndNode in xelRoot)
                {
                    if (nametofind == xndNode.GetAttribute("DocumentName"))
                    {
                        DateTime printtime = Convert.ToDateTime(xndNode.GetAttribute("Time"));
                        if (printtime.AddSeconds(20) < timeofprint)
                        { MessageBox.Show("Passaram mais de 20 segundos desde a impressão de " + nametofind); }
                        else
                        { MessageBox.Show("Passaram menos de 20 segundos desde a impressão de " + nametofind); }
                        contents = xndNode.GetAttribute("DocumentName") + " " + printtime.ToString();
                    }
                    
                    

                    
                    //Your sql insert command will go here;
                }

                MessageBox.Show(contents);
            }
            else { MessageBox.Show("Audit Empty"); }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string schedulePath = @"C:\PrijobAide\Bin\HOLTER.scl";
            if (File.Exists(schedulePath))
            {
                List<string> rawschedule = File.ReadAllLines(schedulePath).ToList();
                string Dayofweek = DateTime.Now.DayOfWeek.ToString();
                string SclTime = "";
                foreach (string s in rawschedule)
                {
                    if (s.Contains(Dayofweek + "="))
                    { SclTime = s.Replace(Dayofweek + "=", ""); }
                }
                if (SclTime != "")
                {
                    if (SclTime.Contains("all"))
                    { MessageBox.Show("ICANPRINT"); }
                    else
                    {
                        if (SclTime.Contains(DateTime.Now.Hour.ToString() + ";"))
                        { MessageBox.Show("ICANPRINT"); }
                    }
                }
                
                
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            Helper internalhelper = new Helper();
            
                
                string file = @"C:\PDFOutput\file0002.pdf";
                string outputPath = @"C:\PDFOutput\";

                ExtractImage(file, outputPath);
                    
                
                
            
        }
        private void ExtractImage(string pdfFile,string path)
        {
            PdfReader pdfReader  = new PdfReader(pdfFile);
            for (int pageNumber = 1; pageNumber <= pdfReader.NumberOfPages; pageNumber++)
            {
                PdfReader pdf = new PdfReader(pdfFile);
                PdfDictionary pg = pdf.GetPageN(pageNumber);
                PdfDictionary res = (PdfDictionary)PdfReader.GetPdfObject(pg.Get(PdfName.RESOURCES));
                PdfDictionary xobj = (PdfDictionary)PdfReader.GetPdfObject(res.Get(PdfName.XOBJECT));
                int i = 0;
                foreach (PdfName name in xobj.Keys)
                {
                    PdfObject obj = xobj.Get(name);
                    if (obj.IsIndirect())
                    {
                        i++;
                        PdfDictionary tg = (PdfDictionary)PdfReader.GetPdfObject(obj);
                        string width = tg.Get(PdfName.WIDTH).ToString();
                        string height = tg.Get(PdfName.HEIGHT).ToString();
                        ImageRenderInfo imgRI = ImageRenderInfo.CreateForXObject(new Matrix(float.Parse(width), float.Parse(height)), (PRIndirectReference)obj, tg);
                        RenderImage(imgRI,path + i.ToString() + ".bmp");
                    }
                }
            }
        }
        private void RenderImage(ImageRenderInfo renderInfo,string imgPath)
        {
            PdfImageObject image = renderInfo.GetImage();
            using (Dotnet dotnetImg = image.GetDrawingImage())
            {
                if (dotnetImg != null)
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        dotnetImg.Save(ms, ImageFormat.Tiff);
                        Bitmap d = new Bitmap(dotnetImg);
                        d.Save(imgPath);
                    }
                }
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            object[] parameters = new object[] { internalMain.towrite, internalMain.gpathPDF, internalMain.gpathECG, internalMain.gpathECGTbl, internalMain.gpathECO, internalMain.gpathECOTbl, internalMain.gpathHOLTER, internalMain.gpathHOLTERTbl, internalMain.gpathMAPA, internalMain.gpathMAPATbl, internalMain.gpathPE, internalMain.gpathPETbl, internalMain.gpathUnconfirmed };
            internalMain.RunPDFOutPut(parameters);
        }

        private void button17_Click(object sender, EventArgs e)
        {
            object[] parameters = new object[] { internalMain.towrite, internalMain.gpathECG, internalMain.gpathECO, internalMain.gpathHOLTER, internalMain.gpathMAPA, internalMain.gpathPE, internalMain.gECGRPrinter, internalMain.gECGIPrinter, internalMain.gECGCPrinter, internalMain.gECORPrinter, internalMain.gECOIPrinter, internalMain.gECOCPrinter, internalMain.gHOLTERRPrinter, internalMain.gHOLTERIPrinter, internalMain.gHOLTERCPrinter, internalMain.gMAPARPrinter, internalMain.gMAPAIPrinter, internalMain.gMAPACPrinter, internalMain.gprintECGR, internalMain.gprintECGI, internalMain.gprintECGC, internalMain.gprintECOR, internalMain.gprintECOI, internalMain.gprintECOC, internalMain.gprintHOLTERR, internalMain.gprintHOLTERI, internalMain.gprintHOLTERC, internalMain.gprintMAPAR, internalMain.gprintMAPAI, internalMain.gprintMAPAC, internalMain.gprintPER, internalMain.gprintPEI, internalMain.gprintPEC, internalMain.gjoinECG, internalMain.gjoinECO, internalMain.gjoinHOLTER, internalMain.gjoinMAPA, internalMain.gjoinPE, internalMain.gBookletPrinter, internalMain.gBookletAudit, internalMain.gBookletProcess, internalMain.gBookletSecsToWait, internalMain.gLastBookletprint, internalMain.gPERPrinter, internalMain.gPEIPrinter, internalMain.gPECPrinter, internalMain.gpathSumatraPDF, internalMain.gStoreReports, internalMain.gStoreImages, internalMain.gStoreComplete, internalMain.gpathUnconfirmed, internalMain.gdebug };
            internalMain.RunFolders(parameters);
        }

        private void button18_Click(object sender, EventArgs e)
        {
            object[] parameters = new object[] { internalMain.towrite, internalMain.gpathPDF, internalMain.gPdfPrinter, internalMain.gEmailInc, internalMain.gpathHOLTERTAdd, internalMain.gpathHOLTERTPass, internalMain.gHOLTERExternalApp, internalMain.gHOLTERExternalAppPath, internalMain.gHOLTERSender, internalMain.gHOLTERReceiver };
            internalMain.RunEmail(parameters);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            string CompletePath = @"C:\Store\Complete";
            List<DirectoryInfo> subDirectories = new List<DirectoryInfo>();
            if (Directory.Exists(CompletePath))
            {
                List <string> di = Directory.GetDirectories (CompletePath).ToList();
                foreach (string dir in di)
                {
                    DirectoryInfo a = new DirectoryInfo(dir);
                    subDirectories.Add(a);
                }
            }
            if (subDirectories.Count > 0)
            {
                List<DirectoryInfo> OrderredsubDirectories = subDirectories.OrderByDescending(d => Convert.ToDateTime(d.Name)).ToList();
                List<string> ProcessedDays = new List<string>();
                List<string> ProcessedId = new List<string>();
                string message = "";
                List<string> Message = new List<string>();
                int name = 30;
                int id = 10;
                foreach (DirectoryInfo sd in OrderredsubDirectories)
                {
                    if (Directory.Exists(sd.FullName))
                    {
                        
                        List<string> Files = Directory.GetFiles(sd.FullName,"HOLTER.C.*.pdf").ToList();
                        if (Files.Count > 0)
                        {
                            foreach (string fl in Files)
                            {
                                List<string> ParnerContent = new List<string>();
                                string partner = fl.Replace(".pdf", ".txt");
                                if (File.Exists(partner))
                                { ParnerContent = File.ReadAllLines(partner).ToList(); }
                                if (ProcessedId.Contains(ParnerContent[1]) == false)
                                {
                                    ProcessedId.Add(ParnerContent[1]);
                                    if (ProcessedDays.Contains(sd.Name) == false)
                                    {
                                        ProcessedDays.Add(sd.Name);
                                        message = message + Environment.NewLine + sd.Name;
                                        Message.Add("");
                                        Message.Add(sd.Name);
                                    }
                                    int missingname = name - ParnerContent[0].Length;
                                    int missingid = id - ParnerContent[1].Length;
                                    string tabname = new String(' ', (missingname));
                                    string tabid = new String(' ', (missingid));
                                    
                                    message = message + Environment.NewLine + ParnerContent[0].Trim() + tabname + ParnerContent[1].Trim() + tabid + " Recebido";
                                 Message.Add(ParnerContent[0].Trim() + tabname + ParnerContent[1].Trim() + tabid + " Recebido");
                                }
                            }
                        }
                    }
                }
                MessageBox.Show(message);
                File.WriteAllLines(@"C:\PriJobAide\Logs\HolterRep.txt",Message.ToArray());
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            object[] parameters = new object[] { internalMain.towrite, internalMain.gpathHOLTERTAdd, internalMain.gpathHOLTERTPass, internalMain.gHOLTERSender, internalMain.gStoreComplete };
            internalMain.RunReport(parameters);
        }

        private void button21_Click(object sender, EventArgs e)
        {
            DirectoryInfo di = new DirectoryInfo(internalMain.gpathHOLTER);
            List<FileInfo> TxtFiles = di.GetFiles("*.txt", SearchOption.TopDirectoryOnly).ToList();
            List<string> Reports = new List<string>();
            List<string> Images = new List<string>();
            List<string> Complete = new List<string>();
            foreach (FileInfo a in TxtFiles)
            {
                if (File.Exists(a.FullName.Replace(".txt", ".pdf")))
                {
                    List<string> content = File.ReadAllLines(a.FullName).ToList();
                    if (content.Count >= 8)
                    {
                        if (content[2] == "R")
                        { Reports.Add(content[0] + "separator" + content[1] + "separator" + content[3] + "separator" + content[4]); }
                        if (content[2] == "I")
                        { Images.Add(content[0] + "separator" + content[1] + "separator" + content[3] + "separator" + content[4]); }
                        if (content[2] == "C")
                        { Complete.Add(content[0] + "separator" + content[1] + "separator" + content[3] + "separator" + content[4]); }
                    }
                }
            }
            List<string> ExistingTrees = new List<string>();
            foreach (string Rep in Reports)
            {
                string[] split;
                try
                {
                     split = Rep.Split(new string[] { "separator" }, StringSplitOptions.None); ;
                }
                catch { split = null; }
                if (split.Length == 4)
                {
                    TreeNode date = new TreeNode(split[2]);
                    date.Name = split[2];
                    
                    if (treeView1.Nodes.ContainsKey(split[2]) == false)
                    { treeView1.Nodes.Add(date); treeView1.Update(); }

                    TreeNode nameandid = new TreeNode(split[0] + " (" + split[1] + ")");
                    nameandid.Name = split[0] + " (" + split[1] + ")";
                    if (treeView1.Nodes[split[2]].Nodes.ContainsKey(split[0] + " (" + split[1] + ")") == false)
                    { treeView1.Nodes[split[2]].Nodes.Add(nameandid); treeView1.Update(); }
                    
                    TreeNode tipology = new TreeNode("R");
                    tipology.Name = "R";
                    if (treeView1.Nodes[split[2]].Nodes[split[0] + " (" + split[1] + ")"].Nodes.ContainsKey("R") == false)
                    { treeView1.Nodes[split[2]].Nodes[split[0] + " (" + split[1] + ")"].Nodes.Add(tipology); treeView1.Update(); }
                    TreeNode state = new TreeNode(split[3]);
                    

                    

                
                }
            }
        }


        
    }
}
