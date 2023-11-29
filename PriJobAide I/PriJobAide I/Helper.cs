using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;
using System.Runtime.InteropServices;
using System.Drawing.Printing;
using System.Diagnostics;
using System.Xml;



namespace PriJobAide
{
    class Helper
    {
        [DllImport("user32.dll")]
        public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("User32.dll")]
        private static extern int ShowWindow(int hwnd, int nCmdShow);
        [DllImport("user32.dll", EntryPoint = "FindWindowEx", CharSet = CharSet.Auto)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(String sClassName, String sAppName);
        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetDefaultPrinter(string Name);
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")]
        static extern byte VkKeyScan(char ch);
        const int WM_KEYDOWN = 0x100;
        [DllImport("user32.dll", EntryPoint = "GetWindowText", CharSet = CharSet.Auto)]
        static extern IntPtr GetWindowCaption(IntPtr hwnd, StringBuilder lpString, int maxCount);
       
        const UInt32 WS_OVERLAPPED = 0;
 const UInt32 WS_POPUP = 0x80000000;
 const UInt32 WS_CHILD = 0x40000000;
 const UInt32 WS_MINIMIZE = 0x20000000;
 const UInt32 WS_VISIBLE = 0x10000000;
 const UInt32 WS_DISABLED = 0x8000000;
 const UInt32 WS_CLIPSIBLINGS = 0x4000000;
 const UInt32 WS_CLIPCHILDREN = 0x2000000;
 const UInt32 WS_MAXIMIZE = 0x1000000;
 const UInt32 WS_CAPTION = 0xC00000;      // WS_BORDER or WS_DLGFRAME  
 const UInt32 WS_BORDER = 0x800000;
 const UInt32 WS_DLGFRAME = 0x400000;
 const UInt32 WS_VSCROLL = 0x200000;
 const UInt32 WS_HSCROLL = 0x100000;
 const UInt32 WS_SYSMENU = 0x80000;
 const UInt32 WS_THICKFRAME = 0x40000;
 const UInt32 WS_GROUP = 0x20000;
 const UInt32 WS_TABSTOP = 0x10000;
 const UInt32 WS_MINIMIZEBOX = 0x20000;
 const UInt32 WS_MAXIMIZEBOX = 0x10000;
 const UInt32 WS_TILED = WS_OVERLAPPED;
 const UInt32 WS_ICONIC = WS_MINIMIZE;
 const UInt32 WS_SIZEBOX = WS_THICKFRAME;
        public bool printerisvalid(string printername)
        {
            bool printerisvalid = false;
            PrinterSettings printer = new PrinterSettings();
            printer.PrinterName = printername;
            printerisvalid = printer.IsValid;
            return printerisvalid;
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
        public string Formatmsg(string msg)
        {
            string retmsg = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.fff") + " : " + msg;
            return retmsg;
        }
        public string TrimIniText(string line)
        {
            try
            {
                string returnstring = "";
                if (String.IsNullOrEmpty(line) != true)
                {
                    if (line.Contains("=") == true)
                    {
                        String[] splitline = line.Split('=');

                        returnstring = splitline[splitline.Length - 1].Trim();
                    }
                    else
                    {
                        returnstring = line.Trim();
                    }

                }

                return returnstring;
            }
            catch
            {
                string failure = "";
                return failure;
            }
        }
        public bool PDFCorrupt(string path)
        {

            try
            {
                PdfSharp.Pdf.PdfDocument doc = PdfSharp.Pdf.IO.PdfReader.Open(path, PdfDocumentOpenMode.Modify);
                doc.Close();
                return false;

            }
            catch
            {

                return true;
            }


        }
        public List<string> ReadPdfFile(string fileName)
        {
            List<string> Out = new List<string>();
            try
            {
                StringBuilder text = new StringBuilder();

                if (File.Exists(fileName))
                {
                    iTextSharp.text.pdf.PdfReader pdfReader = new iTextSharp.text.pdf.PdfReader(fileName);
                    for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                    {
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                        currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                        text.Append(currentText + Environment.NewLine);
                    }
                    pdfReader.Close();
                    string[] lines = text.ToString().Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                    Out = lines.ToList();
                    pdfReader.Dispose();
                }
                return Out;
            }
            catch { return Out; }
        }
        public string PDFFileNameTransformECG(List<string> RawECG)
        {
            string Out = "";

            return Out;
        }
        public void EmailCheck(string user, string pass, string sender, string receiver, string EmailInc)
        {
            try
            {
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                service.Credentials = new WebCredentials(user, pass);
                DateTime cuttoff = DateTime.Now;
                cuttoff = cuttoff.AddDays(-10);
                service.UseDefaultCredentials = false;
                List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
                searchFilterCollection.Add(new SearchFilter.IsGreaterThan(ItemSchema.DateTimeSent, cuttoff));
                searchFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.From, sender));
                searchFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.ToRecipients, receiver));

                searchFilterCollection.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
                SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilterCollection.ToArray());
                service.Url = new Uri(@"https://webmail.tecnifar.pt/ews/exchange.asmx");
                FindItemsResults<Item> findResults = service.FindItems(
                   WellKnownFolderName.Inbox, searchFilter,
                   new ItemView(20));

                foreach (Item item in findResults.Items)
                {
                    if (item is EmailMessage)
                    {

                        EmailMessage message = EmailMessage.Bind(service, new ItemId(item.Id.UniqueId.ToString()), new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments));
                        message.IsRead = true;
                        message.Update(ConflictResolutionMode.AlwaysOverwrite);
                        int s = message.Attachments.Count;


                        foreach (Attachment attachment in message.Attachments)
                        {
                            if (attachment is FileAttachment)
                            {
                                FileAttachment fileAttachment = attachment as FileAttachment;
                                fileAttachment.Load();
                                fileAttachment.Load(EmailInc + @"\" + fileAttachment.Name);

                            }

                        }





                    }
                }
            }
            catch { }
        }
        public void sendKeystrokePrint(IntPtr edit, IntPtr Wpar, IntPtr Lpar)
        {
            PostMessage(edit, 0x111, Wpar, Lpar);
                   
        }
        public void sendKeystrokeEnter(IntPtr edit, IntPtr Wpar, IntPtr Lpar)
        {
            PostMessage(edit, 0x0471, Wpar, Lpar);

        }
        public void endapp(IntPtr edit)
        {




            PostMessage(edit, 0x10, IntPtr.Zero, IntPtr.Zero);
            



        }
        public static bool ContainsAllItems(List<string> pdfcontent, List<string> contenttype)
        {
            int hit = 0;
            foreach (string s in pdfcontent)
            {
                foreach (string f in contenttype)
                {
                    
                    if (s.Contains(f))
                    { hit++; }
                }
            }
            if (hit >= contenttype.Count)
            { return true; }
            else
            { return false; }
            
            

        }
        public int IsECGReport(List<string> pdfcontent, List<string> EcgReportContent)
        {
            
            int sucess = 0;
            int fail = 0;
            if (EcgReportContent.Count > 0)
            {
                try 
                {
                    foreach (string s in EcgReportContent)
                    {
                        string[] split = s.Split('=');
                        if (pdfcontent[Convert.ToInt32(split[0])] == split[1])
                        {
                            sucess++;
                        }
                        else
                        {
                            fail++;
                        }
                    }
                    if ((sucess > 0) && (fail == 0))
                    { return sucess; }
                    else { return 0;}
                }
                catch { return 0; }
            }
            else
            {
                return 0;
            }
            
        }
        public int ExamType(List<string> pdfcontent,List<string> tables)
        {
            int examtype = 0; // zero significa que ainda não se sabe o que é
            try 
            {
                if (pdfcontent.Count > 0)
                {
                   
                    if (tables.Count > 0)
                    {
                        int tblcount = 0;
                        foreach (string tbl in tables)
                        {
                            
                            if (File.Exists(tbl))
                            {
                                
                                if (examtype == 0)
                                {
                                    List<string> ReportContent = new List<string>();
                                    List<string> ImgContent = new List<string>();
                                    List<string> FullTbl = new List<string>();
                                    FullTbl = File.ReadAllLines(tbl).ToList();
                                    int testrepexamtyp = 0;
                                    int testimgexamtyp = 0;
                                    if (FullTbl.Count > 0)
                                    {
                                        foreach (string s in FullTbl)
                                        {
                                            
                                            if (s.Contains("ReportContent="))
                                            {
                                                string linetoadd = s.Replace("ReportContent=", "");
                                                if (linetoadd != "")
                                                {
                                                    ReportContent.Add(s.Replace("ReportContent=", "").Trim());
                                                }
                                            }
                                            if (s.Contains("ImageContent="))
                                            {
                                                string linetoadd = s.Replace("ImageContent=", "");
                                                if (linetoadd != "")
                                                {
                                                    ImgContent.Add(s.Replace("ImageContent=", "").Trim());
                                                }
                                            }
                                            if (s.Contains("ReportExamType="))
                                            {
                                                string linetoadd = s.Replace("ReportExamType=", "");
                                                testrepexamtyp = Convert.ToInt32(linetoadd);
                                            }
                                            if (s.Contains("ImageExamType="))
                                            {
                                                string linetoadd = s.Replace("ImageExamType=", "");
                                                testimgexamtyp = Convert.ToInt32(linetoadd);
                                            }
                                        
                                        }
                                        if (ReportContent.Count > 0)
                                        {
                                            int repsucess = 0;
                                            int repfail = 0;
                                            bool ContainsallRepContent = ContainsAllItems(pdfcontent, ReportContent);


                                            if (ContainsallRepContent == true)
                                                {
                                                    repsucess++;
                                                }
                                                else
                                                {
                                                    repfail++;
                                                }
                                            
                                            if ((repsucess > 0) && (repfail == 0))
                                            { examtype = testrepexamtyp; }
                                            
                                        }
                                        if (ImgContent.Count > 0)
                                        {
                                            int sucess = 0;
                                            int fail = 0;
                                            bool ContainsallImgContent = ContainsAllItems(pdfcontent, ImgContent);
                                            
                                            if (ContainsallImgContent == true)
                                            {
                                                
                                                    sucess++;
                                                }
                                                else
                                                {
                                                    fail++;
                                                
                                            }
                                            if ((sucess > 0) && (fail == 0))
                                            { examtype = testimgexamtyp; }

                                        }
                                    
                                    }
                                }
                                tblcount++;
                            }
                        
                        
                        }
                        if (tblcount == 0)
                        {
                            return 13; //significa que nenhuma tabela existe
                        }
                        else
                        {
                            if (examtype == 0)
                            {
                                return 14;//significa que depois de correr todas as tabelas não sabemos o que o ficheiro é

                            }
                            else
                            {
                                return examtype;
                            }
                        }
                        
                    }
                    else
                    {
                        return 12; //significa que não há tabelas para comparar
                    }
                }
                else
                {
                    return 11;
                }
                
            }
            catch { return 11; }//significa que há um erro neste pdf
        
        }
        public bool isnormalized(string s)
        {
            bool isnormalized = false;
            try
            {
                string[] split = s.Split('.');

                if (split.Length == 6)
                { isnormalized = true; }
                else { isnormalized = false; }
                //MessageBox.Show("0=" + split[0] + " 1=" + split[1] + " 2=" + split[2] + " 3=" + split[3] + " 4=" + split[4] + " 5=" + split[5]);
            }
            catch { isnormalized = false; }

            return isnormalized;
        }
        public string GetName(int ExamType, List<string> pdfcontent, string table)
        {
            string finalname = "default";
            string tempReportNameStart = "";
            string finalReportNameStart = "";
            string tempReportNameEnd = "";
            string finalReportNameEnd = "";
            
            try
            {
                if (pdfcontent.Count > 0)
                {
                    if (File.Exists(table))
                    {
                        List<string> tblcontent = File.ReadAllLines(table).ToList();
                        List<string> RemoveEx = new List<string>();
                        List<string> GetLeft = new List<string>();
                        List<string> GetRight = new List<string>();
                        if (tblcontent.Count > 0)
                        {
                            foreach (string s in tblcontent)
                            {
                                if (ExamType == 1 || ExamType == 3 || ExamType == 5 || ExamType == 7 || ExamType == 9)
                                {
                                    if (s.Contains("ReportnameStart="))
                                    { tempReportNameStart = s.Replace("ReportnameStart=", ""); }
                                    if (s.Contains("ReportnameEnd="))
                                    { tempReportNameEnd = s.Replace("ReportnameEnd=", ""); }
                                    if (s.Contains("ReportnameGetLeft="))
                                    { GetLeft.Add(s.Replace("ReportnameGetLeft=", "")); }
                                    if (s.Contains("ReportnameGetRight="))
                                    { GetRight.Add(s.Replace("ReportnameGetRight=", "")); }
                                    if (s.Contains("ReportnameEx="))
                                    { RemoveEx.Add(s.Replace("ReportnameEx=", "")); }

                                }
                                if (ExamType == 2 || ExamType == 4 || ExamType == 6 || ExamType == 8 || ExamType == 10)
                                {
                                    if (s.Contains("ImagenameStart="))
                                    { tempReportNameStart = s.Replace("ImagenameStart=", ""); }
                                    if (s.Contains("ImagenameEnd="))
                                    { tempReportNameEnd = s.Replace("ImagenameEnd=", ""); }
                                    if (s.Contains("ImagenameGetLeft="))
                                    { GetLeft.Add(s.Replace("ImagenameGetLeft=", "")); }
                                    if (s.Contains("ImagenameGetRight="))
                                    { GetRight.Add(s.Replace("ImagenameGetRight=", "")); }
                                    if (s.Contains("ImagenameEx="))
                                    { RemoveEx.Add(s.Replace("ImagenameEx=", "")); }
                                }
                            }
                            if (tempReportNameStart != "")
                            {
                               
                                bool isNumeric = false;
                                int nameindex = 0;
                                try { nameindex = Convert.ToInt32(tempReportNameStart); isNumeric = true; }
                                catch { isNumeric = false; }
                                if (isNumeric == true)
                                { finalReportNameStart = pdfcontent[nameindex]; }
                                else
                                {
                                    for (int i = 0; i < pdfcontent.Count; i++)
                                    {
                                        if (pdfcontent[i].Contains(tempReportNameStart))
                                        {
                                            nameindex = i;
                                            break;
                                        }
                                    }
                                    finalReportNameStart = pdfcontent[nameindex].Replace(tempReportNameStart, "").Trim();
                                }
                            }
                            else { finalReportNameStart = ""; }
                            if (tempReportNameEnd != "")
                            {
                                bool isNumeric = false;
                                int nameindex = 0;
                                try { nameindex = Convert.ToInt32(tempReportNameEnd); isNumeric = true; }
                                catch { isNumeric = false; }
                                if (isNumeric == true)
                                { finalReportNameEnd = pdfcontent[nameindex]; }
                                else
                                {
                                    for (int i = 0; i < pdfcontent.Count; i++)
                                    {
                                        if (pdfcontent[i].Contains(tempReportNameEnd))
                                        {
                                            nameindex = i - 1;
                                            break;
                                        }
                                    }
                                    finalReportNameEnd = pdfcontent[nameindex].Trim();
                                }
                            }
                            else { finalReportNameEnd = ""; }
                            finalname = finalReportNameStart + " " + finalReportNameEnd;
                            if (GetRight.Count > 0)
                            {
                                foreach (string getRight in GetRight)
                                {
                                    if (getRight != "")
                                    {
                                        string tempfinalname = finalname;
                                        try
                                        {
                                            string[] split = tempfinalname.Split(new string[] { getRight }, StringSplitOptions.None);
                                            finalname = split[1];
                                        }
                                        catch { }
                                    }
                                }
                            }
                            if (GetLeft.Count > 0)
                            {
                                foreach (string getLeft in GetLeft)
                                {
                                    if (getLeft != "")
                                    {
                                        string tempfinalname = finalname;
                                        try
                                        {
                                            string[] split = tempfinalname.Split(new string[] { getLeft }, StringSplitOptions.None);
                                            finalname = split[0];
                                        }
                                        catch { }
                                    }
                                }
                            }
                            if (RemoveEx.Count > 0)
                            {
                                foreach (string removeEx in RemoveEx)
                                {
                                    if (removeEx != "")
                                    {

                                        if (finalname.Contains(removeEx))
                                        { finalname.Replace(removeEx, ""); }
                                    }
                                }
                            }
                            if (finalname == " " || finalname == "")
                            {
                                finalname = "default";
                            }
                        }
                            
                    }                
                }
            }
            catch { }
            return finalname.Trim();
        }
        public string GetId(int ExamType,List<string> pdfcontent,string table)
    {
            string finalid = "default";
            string tempReportIdStart = "";
            string finalReportIdStart = "";
            
         try
            {
                if (pdfcontent.Count > 0)
                {
                    if (File.Exists(table))
                    {
                        List<string> tblcontent = File.ReadAllLines(table).ToList();
                        List<string> RemoveEx = new List<string>();
                        List<string> GetLeft = new List<string>();
                        List<string> GetRight = new List<string>();
                        if (tblcontent.Count > 0)
                        {
                            foreach (string s in tblcontent)
                            {
                                if (ExamType == 1 || ExamType == 3 || ExamType == 5 || ExamType == 7 || ExamType == 9)
                                {
                                    if (s.Contains("ReportidStart="))
                                    { tempReportIdStart = s.Replace("ReportidStart=", ""); }

                                    if (s.Contains("ReportidGetLeft="))
                                    { GetLeft.Add(s.Replace("ReportidGetLeft=", "")); }
                                    if (s.Contains("ReportidGetRight="))
                                    { GetRight.Add(s.Replace("ReportidGetRight=", "")); }
                                    if (s.Contains("ReportidEx="))
                                    { RemoveEx.Add(s.Replace("ReportidEx=", "")); }
                                }
                                if (ExamType == 2 || ExamType == 4 || ExamType == 6 || ExamType == 8 || ExamType == 10)
                                {
                                    if (s.Contains("ImageidStart="))
                                    { tempReportIdStart = s.Replace("ImageidStart=", ""); }

                                    if (s.Contains("ImageidGetLeft="))
                                    { GetLeft.Add(s.Replace("ImageidGetLeft=", "")); }
                                    if (s.Contains("ImageidGetRight="))
                                    { GetRight.Add(s.Replace("ImageidGetRight=", "")); }
                                    if (s.Contains("ImageidEx="))
                                    { RemoveEx.Add(s.Replace("ImageidEx=", "")); }
                                    
                                }
                            }
                            if (tempReportIdStart != "")
                            {
                                bool isNumeric = false;
                                int nameindex = 0;
                                try { nameindex = Convert.ToInt32(tempReportIdStart); isNumeric = true; }
                                catch { isNumeric = false; }
                                if (isNumeric == true)
                                { finalReportIdStart = pdfcontent[nameindex]; }
                                else
                                {
                                    for (int i = 0; i < pdfcontent.Count; i++)
                                    {
                                        if (pdfcontent[i].Contains(tempReportIdStart))
                                        {
                                            nameindex = i;
                                            break;
                                        }
                                    }
                                    string[] split = pdfcontent[nameindex].Split(new string[] { tempReportIdStart }, StringSplitOptions.None);
                                    finalReportIdStart = split[1];
                                }
                            }
                            else { finalReportIdStart = ""; }
                            
                            
                            finalid = finalReportIdStart;
                            if (GetRight.Count > 0)
                            {
                                foreach (string getRight in GetRight)
                                {
                                    if (getRight != "")
                                    {
                                        string tempfinalname = finalid;
                                        try
                                        {
                                            string[] split = tempfinalname.Split(new string[] { getRight }, StringSplitOptions.None);
                                            finalid = split[1];
                                        }
                                        catch { }
                                    }
                                }
                            }
                            if (GetLeft.Count > 0)
                            {
                                foreach (string getLeft in GetLeft)
                                {
                                    if (getLeft != "")
                                    {
                                        string tempfinalname = finalid;
                                        try
                                        {
                                            string[] split = tempfinalname.Split(new string[] { getLeft }, StringSplitOptions.None);
                                            finalid = split[0];
                                        }
                                        catch { }
                                    }
                                }
                            }
                            if (RemoveEx.Count > 0)
                            {
                                foreach (string removeEx in RemoveEx)
                                {
                                    if (removeEx != "")
                                    {

                                        if (finalid.Contains(removeEx))
                                        { finalid.Replace(removeEx, ""); }
                                    }
                                }
                            }
                            if (finalid == " " || finalid == "")
                            {
                                finalid = "default";
                            }
                        }
                            
                    }                
                }
            }
            catch { }
         return finalid.Trim();
    }
        public string GetDate(int ExamType, List<string> pdfcontent, string table)
    {
        string finalDate = "default";
        string tempReportDateStart = "";
        string finalReportDateStart = "";

        try
        {
            if (pdfcontent.Count > 0)
            {
                if (File.Exists(table))
                {
                    List<string> tblcontent = File.ReadAllLines(table).ToList();
                    List<string> RemoveEx = new List<string>();
                    List<string> GetLeft = new List<string>();
                    List<string> GetRight = new List<string>();
                    if (tblcontent.Count > 0)
                    {
                        foreach (string s in tblcontent)
                        {
                            if (ExamType == 1 || ExamType == 3 || ExamType == 5 || ExamType == 7 || ExamType == 9)
                            {
                                if (s.Contains("ReportDateStart="))
                                { tempReportDateStart = s.Replace("ReportDateStart=", ""); }
                                if (s.Contains("ReportDateGetLeft="))
                                { GetLeft.Add(s.Replace("ReportDateGetLeft=", "")); }
                                if (s.Contains("ReportDateGetRight="))
                                { GetRight.Add(s.Replace("ReportDateGetRight=", "")); }
                                if (s.Contains("ReportDateEx="))
                                { RemoveEx.Add(s.Replace("ReportDateEx=", "")); }

                            }
                            if (ExamType == 2 || ExamType == 4 || ExamType == 6 || ExamType == 8 || ExamType == 10)
                            {
                                if (s.Contains("ImageDateStart="))
                                { tempReportDateStart = s.Replace("ImageDateStart=", ""); }
                                if (s.Contains("ImageDateGetLeft="))
                                { GetLeft.Add(s.Replace("ImageDateGetLeft=", "")); }
                                if (s.Contains("ImageDateGetRight="))
                                { GetRight.Add(s.Replace("ImageDateGetRight=", "")); }
                                if (s.Contains("ImageDateEx="))
                                { RemoveEx.Add(s.Replace("ImageDateEx=", "")); }

                            }
                        }
                        if (tempReportDateStart != "")
                        {
                            bool isNumeric = false;
                            int nameindex = 0;
                            try { nameindex = Convert.ToInt32(tempReportDateStart); isNumeric = true; }
                            catch { isNumeric = false; }
                            if (isNumeric == true)
                            { finalReportDateStart = pdfcontent[nameindex]; }
                            else
                            {
                                for (int i = 0; i < pdfcontent.Count; i++)
                                {
                                    if (pdfcontent[i].Contains(tempReportDateStart))
                                    {
                                        nameindex = i;
                                        break;
                                    }
                                }
                                string[] split = pdfcontent[nameindex].Split(new string[] { tempReportDateStart }, StringSplitOptions.None);
                                finalReportDateStart = split[1];
                            }
                        }
                        else { finalReportDateStart = ""; }


                        finalDate = finalReportDateStart;
                        
                        if (GetRight.Count > 0)
                        {
                            foreach (string getRight in GetRight)
                            {
                                if (getRight != "")
                                {
                                    string tempfinalname = finalDate;
                                    try
                                    {
                                        string[] split = tempfinalname.Split(new string[] { getRight }, StringSplitOptions.None);
                                        finalDate = split[1];
                                    }
                                    catch { }
                                }
                            }
                        }
                        if (GetLeft.Count > 0)
                        {
                            foreach (string getLeft in GetLeft)
                            {
                                if (getLeft != "")
                                {
                                    string tempfinalname = finalDate;
                                    try
                                    {
                                        string[] split = tempfinalname.Split(new string[] { getLeft }, StringSplitOptions.None);
                                        finalDate = split[0];
                                    }
                                    catch { }
                                }
                            }
                        }
                        if (RemoveEx.Count > 0)
                        {
                            foreach (string removeEx in RemoveEx)
                            {
                                if (removeEx != "")
                                {

                                    if (finalDate.Contains(removeEx))
                                    { finalDate.Replace(removeEx, ""); }
                                }
                            }
                        }
                        finalDate = finalDate.Trim();
                        if (finalDate.Length > 11)
                        {
                            finalDate = finalDate.Substring(0, 10);
                        }

                        if (finalDate == " " || finalDate == "")
                        {
                            finalDate = "default";
                        }
                    }

                }
            }
        }
        catch { }
        return finalDate.Trim();
    }
        public string GetNormalizedName(string incoming)
    {

        string final = "default";
        string firstname = "";
        string lastname = "";
        incoming = incoming.Trim();
        if (incoming != "")
        {
            if (incoming.Contains(',')) 
            {
                try
                {
                    string[] split = incoming.Split(',');
                    incoming = split[1].Trim() + " " + split[0].Trim();

                }
                catch {  }
            }
                if (incoming.Contains(' '))
                {
                    
                    try
                    {
                        string[] split2 = incoming.Split(' ');
                        
                        if (split2.Length > 2)
                        {
                            if (split2[0].Length < 3 || split2[0].Contains("Maria") || split2[0].Contains("MARIA"))
                            {
                                firstname = split2[0].Trim() + " " + split2[1].Trim();
                                lastname = split2[split2.Length - 1].Trim();
                                final = firstname + " " + lastname;
                                
                            }
                            else
                            {
                                firstname = split2[0].Trim();
                                lastname = split2[split2.Length - 1].Trim();
                                final = firstname + " " + lastname;
                            }
                        }
                        else
                        {
                            firstname = split2[0];
                            lastname = split2[1];
                            final = firstname + " " + lastname;
                        }

                    }
                    catch {  }
                }
            
            
        }
        
        return final.ToUpper();
    }
        public string GetNormalizedID(string original)
    {
        string result = "default";
        try { result = new string(original.Where(c => Char.IsDigit(c)).ToArray()); }
        catch { }
        return result;
    }
        public bool isProcessRunning(string pname)
    {
        bool isRunning = true;
        try
        {
            Process[] name = Process.GetProcessesByName(pname);
            if (name.Length == 0)
            {
                Thread.Sleep(5000);
                Process[] nameretry = Process.GetProcessesByName(pname);
                if (nameretry.Length == 0)
                {
                    isRunning = false;
                }
                else { isRunning = true; }
            }
            else
            {
                isRunning = true;
            }
        }
        catch { isRunning = true; }
        return isRunning;
    }
        public bool FoundPrintJob(string Audits, string jobname, DateTime JobStartTime, int secstowait)
    {
        bool FoundJob = false;
        bool FoundFile = false;
        DateTime startcheck = DateTime.Now;
        DirectoryInfo Audit = new DirectoryInfo(Audits);
        
        
            string AuditIdentifier = "Print-in." + DateTime.Now.ToString("yyyyMMdd") + ".xml";
            if (File.Exists(Audits + @"\" + AuditIdentifier) == false)
            {
                
                DateTime starcheckone = DateTime.Now;
                
                while (File.Exists(Audits + @"\" + AuditIdentifier) == false)
                {
                    
                    if (DateTime.Now > starcheckone.AddMinutes(3))
                    { break;  }
                    Thread.Sleep(1000);
                    
                }
            }
            FoundFile = File.Exists(Audits + @"\" + AuditIdentifier);
            
            if (FoundFile == true)
            {
                while
                    (FoundJob == false)
                {
                    
                    Audit.Refresh();
                    
                    
                        
                        
                        XmlDocument xdcDocument = new XmlDocument();
                        xdcDocument.Load(Audits + @"\" + AuditIdentifier);
                        
                        XmlElement xelRoot = xdcDocument.DocumentElement;
                        foreach (XmlElement xndNode in xelRoot)
                        {
                            
                               
                            if (xndNode.GetAttribute("DocumentName").Contains(jobname))
                            {
                                DateTime printtime = Convert.ToDateTime(xndNode.GetAttribute("Time"));
                                
                                if (printtime > JobStartTime)
                                {
                                    
                                    FoundJob = true; break;
                                }
                            }
                        }
                    
                    if (DateTime.Now > startcheck.AddMinutes(3))
                    { FoundJob = false; break; }
                    Thread.Sleep(2000);
                }
            }
        return FoundJob;
    }
        public DateTime LastBookletPrintTime(string audits)
    {
        DateTime LastBookletPrintTime = DateTime.Now;
        XmlDocument xdcDocument = new XmlDocument();

        DirectoryInfo Audit = new DirectoryInfo(audits);
        List<FileInfo> AuditList = Audit.GetFiles("*.xml").Where(x => x.LastWriteTime.Date == DateTime.Today.Date).ToList();
        List<FileInfo> OrderedAuditList = AuditList.OrderByDescending(x => x.LastWriteTime).ToList();
        if (OrderedAuditList.Count > 0)
        {
            OrderedAuditList[0].Refresh();
            xdcDocument.Load(OrderedAuditList[0].FullName);

            XmlElement xelRoot = xdcDocument.DocumentElement;


            
            foreach (XmlElement xndNode in xelRoot)
            {

                LastBookletPrintTime = Convert.ToDateTime(xndNode.GetAttribute("Time"));

                
            }

           
        }
        
        return LastBookletPrintTime;
    }
        public bool MergePDF(string outputFilePath, string[] pdfFiles)
    {
        bool sucess = false;
        try
        {
            int error = 0;
            foreach (string pdfFile in pdfFiles)
            {
                if (this.FileInUse(pdfFile) == true)

                { error++; }
            }
            if (error == 0)
            {
                PdfSharp.Pdf.PdfDocument outputPDFDocument = new PdfSharp.Pdf.PdfDocument();
                foreach (string pdfFile in pdfFiles)
                {

                    PdfSharp.Pdf.PdfDocument inputPDFDocument = PdfSharp.Pdf.IO.PdfReader.Open(pdfFile, PdfDocumentOpenMode.Import);
                    outputPDFDocument.Version = inputPDFDocument.Version;
                    foreach (PdfSharp.Pdf.PdfPage page in inputPDFDocument.Pages)
                    {
                        outputPDFDocument.AddPage(page);
                    }
                    
                }
                if (File.Exists(outputFilePath))
                {
                    File.Delete(outputFilePath);
                }
               
                outputPDFDocument.Save(outputFilePath);
                outputPDFDocument.Dispose();
               
                sucess = true;

            }
            else { sucess = false; }
        }
        catch
        {
            sucess = false;
        }
        return sucess;
    }
        public bool CanIPrintHolters()
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
                    {  SclTime = s.Replace(Dayofweek + "=", ""); break; }
                }
                if (SclTime != "")
                {
                    if (SclTime.Contains("all"))
                    { return true; }
                    else
                    {
                        if (SclTime.Contains(DateTime.Now.Hour.ToString() + ";"))
                        { return true; }
                        else { return false; }
                    }
                }
                else { return false; }


            }
            else { return false; }
            
        }
        public void PrintRPS(string filename, string ExecPath,string PrinterName, string identifierPath)
        {
            
            string startDefaultPrinter = "";
            PrinterSettings settings = new PrinterSettings();
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                settings.PrinterName = printer;
                if (settings.IsDefaultPrinter)
                { startDefaultPrinter = settings.PrinterName; break; }
            }
            if (startDefaultPrinter != PrinterName)
            { SetDefaultPrinter(PrinterName); }
            
            
            Thread.Sleep(3000);
            IntPtr hwnd = IntPtr.Zero;
            IntPtr hwndchild = IntPtr.Zero;
            IntPtr hwndchild2 = IntPtr.Zero;
            var p = new System.Diagnostics.Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            
            
            DateTime StarCheckOne = DateTime.Now;
            #region UnsuccessFull
            /* try
            {
                while (p.MainWindowHandle == IntPtr.Zero)
                {
                    // Discard cached information about the process
                    // because MainWindowHandle might be cached.
                    if (FindWindow("#32770", "Vision Premier") != IntPtr.Zero)
                    {
                        hwndchild = FindWindow("#32770", "Vision Premier");
                        List<IntPtr> ChildWindows = this.GetAllChildrenWindowHandles(hwndchild, 10);
                        foreach (IntPtr a in ChildWindows)
                        {
                            string caption = this.GetWindowCaption(a);
                            if (caption == "Erro na abertura do relatório.")
                            { File.WriteAllText(filename.Replace(".RPS", ".txt"), "Ficheiro danificado."); break; }
                            
                        }
                        PostMessage(hwndchild, 0x111, (IntPtr)0x00000002, (IntPtr)0x0010073E);
                        break;
                    }
                    p.Refresh();

                    Thread.Sleep(10);
                    if (StarCheckOne.AddMinutes(4) < DateTime.Now)
                    { break; }
                }

                bool printed = false;
                DateTime Starcheck = DateTime.Now;
                while (p.MainWindowHandle != IntPtr.Zero)
                {
                    p.Refresh();
                    hwnd = p.MainWindowHandle;
                    Thread.Sleep(100);
                    if (Starcheck.AddMinutes(5) < DateTime.Now)
                    {
                        
                        break;
                    }
                    int style = GetWindowLong(hwnd, -16);
                    if ((style & WS_MAXIMIZE) == WS_MAXIMIZE)
                    {
                        ShowWindow((int)hwnd, 2);
                        Thread.Sleep(100);
                        //It's maximized
                    }
                        
                        
                        
                    PostMessage(hwnd, 0x111, (IntPtr)0x0001835F, (IntPtr)0x00000000);
                    PostMessage(hwnd, 0x111, (IntPtr)0x0001835F, (IntPtr)0x00000000);// abrir Imprimir
                    
                    Thread.Sleep(10);
                    if (FindWindow("#32770", "Imprimir") != IntPtr.Zero && printed == false)
                    {
                        while (FindWindow("#32770", "Imprimir") != IntPtr.Zero )
                        {
                            
                            hwndchild2 = FindWindow("#32770", "Imprimir");

                            ShowWindow((int)hwndchild2, 2);
                            Thread.Sleep(1000);
                            //It's maximized
                            PostMessage(hwndchild2, WM_KEYDOWN, (IntPtr)VkKeyScan((char)13), IntPtr.Zero);
                            File.WriteAllText(filename.Replace(".RPS", ".txt"), "Impresso.");
                            printed = true;
                            Thread.Sleep(1000);
                            break;

                        }
                       
                    }
                    Thread.Sleep(1000);
                    this.endapp(hwnd);
                    p.WaitForExit();
                    
                }
                

            }
            catch
            {
                // The process has probably exited,
                // so accessing MainWindowHandle threw an exception
            }*/

            #endregion
            if (File.Exists(identifierPath))
            {
                List<string> State = File.ReadAllLines(identifierPath).ToList();
            if (State.Count >= 1)
            {
                while ((State[0] != "Finished." && State[0] != "Ficheiro danificado."))
                {
                    if (StarCheckOne.AddMinutes(1) < DateTime.Now)
                    { break; }
                    State = File.ReadAllLines(identifierPath).ToList();
                    if (State[0] == "Finished." || State[0] == "Ficheiro danificado.")
                    { break; }
                    
                    if (State[0] == "Start")
                    {
                        p.StartInfo.FileName = ExecPath;
                        p.StartInfo.CreateNoWindow = false;
                        p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                        p.StartInfo.Arguments = filename;
                        p.StartInfo.RedirectStandardOutput = false;
                        p.StartInfo.RedirectStandardError = true;
                        p.StartInfo.UseShellExecute = false;
                        p.Start();
                        File.WriteAllText(identifierPath, "Started");
                    }
                    if (State[0] == "Started")
                    {
                        p.Refresh();
                        Thread.Sleep(100);
                        if (p.MainWindowHandle == IntPtr.Zero)
                        {
                            
                                // Discard cached information about the process
                                // because MainWindowHandle might be cached.
                                if (FindWindow("#32770", "Vision Premier") != IntPtr.Zero)
                                {
                                    hwndchild = FindWindow("#32770", "Vision Premier");
                                    List<IntPtr> ChildWindows = this.GetAllChildrenWindowHandles(hwndchild, 10);
                                    foreach (IntPtr a in ChildWindows)
                                    {
                                        string caption = this.GetWindowCaption(a);
                                        if (caption == "Erro na abertura do relatório.")
                                        { File.WriteAllText(identifierPath, "Ficheiro danificado."); break; }

                                    }
                                    PostMessage(hwndchild, 0x111, (IntPtr)0x00000002, (IntPtr)0x0010073E);
                                    State = File.ReadAllLines(identifierPath).ToList();
                                    if (State[0] != "Ficheiro danificado.")
                                    {
                                        File.WriteAllText(identifierPath, "Cleared");
                                    }
                                    break;
                                }
                                p.Refresh();

                                Thread.Sleep(10);
                               
                            
                        }
                        else
                        {
                            if (State[0] != "Ficheiro danificado." || State[0] != "Cleared" || State[0] != "Impresso.")
                            {
                                File.WriteAllText(identifierPath, "Cleared");
                            }
                        }
                    }
                    if (State[0] == "Cleared")
                    {
                        p.Refresh();
                        Thread.Sleep(100);
                        if (p.MainWindowHandle != IntPtr.Zero)
                        {
                            
                            hwnd = p.MainWindowHandle;
                            int style = GetWindowLong(hwnd, -16);
                            if ((style & WS_MAXIMIZE) == WS_MAXIMIZE)
                            {
                                ShowWindow((int)hwnd, 2);
                                Thread.Sleep(100);
                                //It's maximized
                            }
                            if (FindWindow("#32770", "Imprimir") == IntPtr.Zero)
                            {
                                PostMessage(hwnd, 0x111, (IntPtr)0x0001835F, (IntPtr)0x00000000);
                            }
                            else
                            {
                                hwndchild2 = FindWindow("#32770", "Imprimir");

                                ShowWindow((int)hwndchild2, 2);
                                Thread.Sleep(1000);

                                while (FindWindow("#32770", "Imprimir") != IntPtr.Zero)
                                {
                                    

                                    p.Refresh();
                                    PostMessage(hwndchild2, WM_KEYDOWN, (IntPtr)VkKeyScan((char)13), IntPtr.Zero);
                                    File.WriteAllText(identifierPath, "Impresso.");

                                    
                                    State = File.ReadAllLines(identifierPath).ToList();
                                    
                                }
                            }
                        
                        }
                    }
                    if (State[0] == "Impresso.")
                    {
                        p.Refresh();
                        Thread.Sleep(100);
                        if (p.MainWindowHandle != IntPtr.Zero)
                        {

                            hwnd = p.MainWindowHandle;
                            if (FindWindow("#32770", "Imprimir") == IntPtr.Zero)
                            {
                                while (p.MainWindowHandle != IntPtr.Zero)
                                {
                                    Thread.Sleep(100);
                                    p.Refresh();
                                    State = File.ReadAllLines(identifierPath).ToList();
                                    if (State[0] != "Finished.")
                                    {
                                        PostMessage(hwnd, 0x10, IntPtr.Zero, IntPtr.Zero);
                                        File.WriteAllText(identifierPath, "Finished.");
                                        State = File.ReadAllLines(identifierPath).ToList();
                                    }
                                    else { break; }
                                    if (p.HasExited == true)
                                    { break; }
                                    
                                }
                            }
                        }
                    }
                    
                }
                SetDefaultPrinter(startDefaultPrinter);
                
            }
            }
        }
        public List<IntPtr> GetAllChildrenWindowHandles(IntPtr hParent, int maxCount)
        {
            List<IntPtr> result = new List<IntPtr>();
            int ct = 0;
            IntPtr prevChild = IntPtr.Zero;
            IntPtr currChild = IntPtr.Zero;
            while (true && ct < maxCount)
            {
                currChild = FindWindowEx(hParent, prevChild, null, null);
                if (currChild == IntPtr.Zero) break;
                result.Add(currChild);
                prevChild = currChild;
                ++ct;
            }
            return result;
        }
        public string GetWindowCaption(IntPtr hwnd)
        {
            StringBuilder sb = new StringBuilder(256);
            GetWindowCaption(hwnd, sb, 256);
            return sb.ToString();
        }
        

    }

    class Prijob
    { 
    private string Iname;
    private string Iid;
    private List<List<string>> Reports;
    private List<List<string>> Images;
    private List<List<string>> Complete;
    public Prijob(string name, string id)
    {
        Iname = name;
        Iid = id;
        Reports = new List<List<string>>();
        Images = new List<List<string>>();
        Complete = new List<List<string>>();
    }
    public List<List<string>> GReports
    {
        get
        {
            return Reports;
        }
    }
    public List<List<string>> GImages
    { get { return Images; } }
    public List<List<string>> GComplete
    {

        get { return Complete; }
    
    } 
    public string Gname
    { get { return Iname; } }
    public string GId
    { get { return Iid; } }
    public void AddReport(List<string> report)
    { Reports.Add(report); }
    public void AddImage(List<string> image)
    { Images.Add(image); }
    public void AddComplete(List<string> complete)
    { Complete.Add(complete); }
    }
}
