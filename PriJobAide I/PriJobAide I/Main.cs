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
using System.Diagnostics;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;


namespace PriJobAide
{
    public partial class Main : Form
    {
        public ListLogger towrite = new ListLogger();
        Helper MyHelper = new Helper();
        FrmDebug frmdebug;
        FrmSetup frmsetup;

        #region Globsl Vsrisbels
        private bool firstload;
        public bool Gfirstload
        { get { return firstload; } set { firstload = value; } }
        private int state;
        private string PathIni = @"C:\PriJobAide\PriJobAide.ini";
        public int gstate
        {
            get { return state; }
            set { state = value; }
        }
        private string PathLogs = @"C:\PriJobAide\Logs";
        public string gPathLogs
        {
            get { return PathLogs; }
            set { PathLogs = value; }
        }
        private string Local;
        public string gLocal
        {
            get { return Local; }
            set { Local = value; }
        }
        private string PdfPrinter;
        public string gPdfPrinter
        {
            get { return PdfPrinter; }
            set { PdfPrinter = value; }
        }
        private int DaystoFlush;
        public int gDaystoFlush
        {
            get { return DaystoFlush; }
            set { DaystoFlush = value; }
        }
        public string gPathIni
        {
            get { return PathIni; }
            set { PathIni = value; }
        }
        private bool debug;
        public bool gdebug
        {
            get { return debug; }
            set { debug = value; }
        }
        private DateTime lastinicheck;
        public DateTime glastinicheck
        {

            get { return lastinicheck; }
            set { lastinicheck = value; }
        }
        private string imediate;
        public string gimediate
        {
            get { return imediate; }
            set { imediate = value; }
        }
        private string pathSumatraPDF;
        public string gpathSumatraPDF
        { get { return pathSumatraPDF; } set { pathSumatraPDF = value; } }
        private string pathECG;
        public string gpathECG
        { get { return pathECG; } set { pathECG = value; } }
        private bool printECGR;
        public bool gprintECGR
        { get { return printECGR; } set { printECGR = value; } }
        private string ECGRPrinter;
        public string gECGRPrinter
        { get { return ECGRPrinter; } set { ECGRPrinter = value; } }
        private bool printECGI;
        public bool gprintECGI
        { get { return printECGI; } set { printECGI = value; } }
        private string ECGIPrinter;
        public string gECGIPrinter
        { get { return ECGIPrinter; } set { ECGIPrinter = value; } }
        private bool printECGC;
        public bool gprintECGC
        { get { return printECGC; } set { printECGC = value; } }
        private string ECGCPrinter;
        public string gECGCPrinter
        { get { return ECGCPrinter; } set { ECGCPrinter = value; } }
        private bool joinECG;
        public bool gjoinECG
        { get { return joinECG; } set { joinECG = value; } }
        private string pathECGTbl;
        public string gpathECGTbl
        { get { return pathECGTbl; } set { pathECGTbl = value; } }
        //endECG
        //Start ECO
        private string pathECO;
        public string gpathECO
        { get { return pathECO; } set { pathECO = value; } }
        private bool printECOR;
        public bool gprintECOR
        { get { return printECOR; } set { printECOR = value; } }
        private string ECORPrinter;
        public string gECORPrinter
        { get { return ECORPrinter; } set { ECORPrinter = value; } }
        private bool printECOI;
        public bool gprintECOI
        { get { return printECOI; } set { printECOI = value; } }
        private string ECOIPrinter;
        public string gECOIPrinter
        { get { return ECOIPrinter; } set { ECOIPrinter = value; } }
        private bool printECOC;
        public bool gprintECOC
        { get { return printECOC; } set { printECOC = value; } }
        private string ECOCPrinter;
        public string gECOCPrinter
        { get { return ECOCPrinter; } set { ECOCPrinter = value; } }
        private bool joinECO;
        public bool gjoinECO
        { get { return joinECO; } set { joinECO = value; } }
        private string pathECOTbl;
        public string gpathECOTbl
        { get { return pathECOTbl; } set { pathECOTbl = value; } }
        //end ECO
        //star Holter
        private string pathHOLTER;
        public string gpathHOLTER
        { get { return pathHOLTER; } set { pathHOLTER = value; } }
        private bool printHOLTERR;
        public bool gprintHOLTERR
        { get { return printHOLTERR; } set { printHOLTERR = value; } }
        private string HOLTERRPrinter;
        public string gHOLTERRPrinter
        { get { return HOLTERRPrinter; } set { HOLTERRPrinter = value; } }
        private bool printHOLTERI;
        public bool gprintHOLTERI
        { get { return printHOLTERI; } set { printHOLTERI = value; } }
        private string HOLTERIPrinter;
        public string gHOLTERIPrinter
        { get { return HOLTERIPrinter; } set { HOLTERIPrinter = value; } }
        private bool printHOLTERC;
        public bool gprintHOLTERC
        { get { return printHOLTERC; } set { printHOLTERC = value; } }
        private string HOLTERCPrinter;
        public string gHOLTERCPrinter
        { get { return HOLTERCPrinter; } set { HOLTERCPrinter = value; } }
        private bool joinHOLTER;
        public bool gjoinHOLTER
        { get { return joinHOLTER; } set { joinHOLTER = value; } }
        private string pathHOLTERTbl;
        public string gpathHOLTERTbl
        { get { return pathHOLTERTbl; } set { pathHOLTERTbl = value; } }
        private string pathHOLTERAdd;
        public string gpathHOLTERTAdd
        { get { return pathHOLTERAdd; } set { pathHOLTERAdd = value; } }
        private string pathHOLTERPass;
        public string gpathHOLTERTPass
        { get { return pathHOLTERPass; } set { pathHOLTERPass = value; } }
        private string HOLTERSender;
        public string gHOLTERSender
        { get { return HOLTERSender; } set { HOLTERSender = value; } }
        private string HOLTERReceiver;
        public string gHOLTERReceiver
        { get { return HOLTERReceiver; } set { HOLTERReceiver = value; } }
        private string HOLTERExternalApp;
        public string gHOLTERExternalApp
        { get { return HOLTERExternalApp; } set { HOLTERExternalApp = value; } }
        private string HOLTERExternalAppPath;
        public string gHOLTERExternalAppPath
        { get { return HOLTERExternalAppPath; } set { HOLTERExternalAppPath = value; } }
        private string StoreReports;
        public string gStoreReports
        { get { return StoreReports; } set { StoreReports = value; } }
        private string StoreImages;
        public string gStoreImages
        { get { return StoreImages; } set { StoreImages = value; } }
        private string StoreComplete;
        public string gStoreComplete
        { get { return StoreComplete; } set { StoreComplete = value; } }

        //End HOLTER
        private string pathUnconfirmed;
        public string gpathUnconfirmed
        { get { return pathUnconfirmed; } set { pathUnconfirmed = value; } }
        // Start MAPA
        private string pathMAPA;
        public string gpathMAPA
        { get { return pathMAPA; } set { pathMAPA = value; } }
        private bool printMAPAR;
        public bool gprintMAPAR
        { get { return printMAPAR; } set { printMAPAR = value; } }
        private string MAPARPrinter;
        public string gMAPARPrinter
        { get { return MAPARPrinter; } set { MAPARPrinter = value; } }
        private bool printMAPAI;
        public bool gprintMAPAI
        { get { return printMAPAI; } set { printMAPAI = value; } }
        private string MAPAIPrinter;
        public string gMAPAIPrinter
        { get { return MAPAIPrinter; } set { MAPAIPrinter = value; } }
        private bool printMAPAC;
        public bool gprintMAPAC
        { get { return printMAPAC; } set { printMAPAC = value; } }
        private string MAPACPrinter;
        public string gMAPACPrinter
        { get { return MAPACPrinter; } set { MAPACPrinter = value; } }
        private bool joinMAPA;
        public bool gjoinMAPA
        { get { return joinMAPA; } set { joinMAPA = value; } }
        private string pathMAPATbl;
        public string gpathMAPATbl
        { get { return pathMAPATbl; } set { pathMAPATbl = value; } }
        // END MAPA
        //START PE
        private string pathPE;
        public string gpathPE
        { get { return pathPE; } set { pathPE = value; } }
        private bool printPER;
        public bool gprintPER
        { get { return printPER; } set { printPER = value; } }
        private string PERPrinter;
        public string gPERPrinter
        { get { return PERPrinter; } set { PERPrinter = value; } }
        private bool printPEI;
        public bool gprintPEI
        { get { return printPEI; } set { printPEI = value; } }
        private string PEIPrinter;
        public string gPEIPrinter
        { get { return PEIPrinter; } set { PEIPrinter = value; } }
        private bool printPEC;
        public bool gprintPEC
        { get { return printPEC; } set { printPEC = value; } }
        private string PECPrinter;
        public string gPECPrinter
        { get { return PECPrinter; } set { PECPrinter = value; } }
        private bool joinPE;
        public bool gjoinPE
        { get { return joinPE; } set { joinPE = value; } }
        private string pathPETbl;
        public string gpathPETbl
        { get { return pathPETbl; } set { pathPETbl = value; } }
        private string pathPDF;
        public string gpathPDF
        {
            get { return pathPDF; }
            set { pathPDF = value; }

        }
        private string BookletPrinter;
        public string gBookletPrinter
        { get { return BookletPrinter; } set { BookletPrinter = value; } }
        private string BookletAudit;
        public string gBookletAudit
        { get { return BookletAudit; } set { BookletAudit = value; } }
        private int BookletSecsToWait;
        public int gBookletSecsToWait
        {
            get { return BookletSecsToWait; }
            set { BookletSecsToWait = value; }
        }
        private string BookletProcess;
        public string gBookletProcess
        { get { return BookletProcess; } set { BookletProcess = value; } }
        private DateTime LastBookletprint;
        public DateTime gLastBookletprint
        {
            get { return LastBookletprint; }
            set { LastBookletprint = value; }
        }
        private DateTime LastFlush;
        public DateTime gLastFlush
        {
            get { return LastFlush; }
            set { LastFlush = value; }
        }
        private string EmailInc;
        public string gEmailInc
        { get { return EmailInc; } set { EmailInc = value; } }
        #endregion
        public  List<TreeNode> GetAllNodes(TreeView _self)
        {
            List<TreeNode> result = new List<TreeNode>();
            foreach (TreeNode child in _self.Nodes)
            {
                result.AddRange(GetAllNodes(child));
            }
            return result;
        }
        public  List<TreeNode> GetAllNodes(TreeNode _self)
        {
            List<TreeNode> result = new List<TreeNode>();
            result.Add(_self);
            foreach (TreeNode child in _self.Nodes)
            {
                result.AddRange(GetAllNodes(child));
            }
            return result;
        }
        public Main(int newstate , bool newdebug)
        {
            gdebug = newdebug;
            gstate = newstate;
            lastinicheck = DateTime.MinValue;
            frmdebug = new FrmDebug(this);
            Gfirstload = true;
            if (newdebug == true)
            {
                frmdebug.Show();
            }
            else
            { frmdebug.Hide(); }
            frmsetup = new FrmSetup();
            InitializeComponent();
        }
        private void Main_Load(object sender, EventArgs e)
        {
            gLastFlush = DateTime.MinValue;
            glastinicheck = DateTime.MinValue;
        }
        //Timers
        private void tmrupdate_Tick(object sender, EventArgs e)
        {
            #region UPdateFrmDebug
            bool testdebug = this.gdebug;
            if (testdebug == true)
            {
                if (frmdebug.Visible == true)
                {

                }
                else
                {
                    try
                    {
                        frmdebug.Show();
                    }
                    catch
                    {
                        frmdebug = new FrmDebug(this);
                        frmdebug.Show();
                    }
                }
            }
            else
            {
                if (frmdebug.Visible == true)
                {
                    frmdebug.Hide();
                }
                else
                {

                }
            } 
            #endregion
            #region FrmSetup
            if (frmsetup.Visible == true)
            {
                this.Hide();
            } 
            #endregion
            if (BCKINIReader.IsBusy == false)
            {
                btnIni.BackColor = Color.LightGreen;
                object[] parameters = new object[] { gPathIni, glastinicheck, towrite };
                BCKINIReader.RunWorkerAsync(parameters);
            }
            else
            { btnIni.BackColor = Color.Yellow; }
            txtimediate.Text = gimediate;
            if (timerFlush.Enabled == true)
            {
                if (BCKFlush.IsBusy)
                { btnFlush.BackColor = Color.LightGreen; }
                else { btnFlush.BackColor = Color.Yellow; }
                btnFolders.BackColor = Color.IndianRed;
                btnPDF.BackColor = Color.IndianRed;
                
            }
            else { btnFlush.BackColor = Color.IndianRed; }
            if (timerFolders.Enabled == true)
            {
                if (BCKPDFOutput.IsBusy)
                { btnPDF.BackColor = Color.LightGreen; }
                else { btnPDF.BackColor = Color.Yellow; }
                if (BCKFolders.IsBusy)
                { btnFolders.BackColor = Color.LightGreen; }
                else { btnFolders.BackColor = Color.Yellow; }
            }
            else { btnPDF.BackColor = Color.IndianRed; btnFolders.BackColor = Color.IndianRed; }
            if (BCKUpdateLst.IsBusy)
            { btnUplst.BackColor = Color.LightGreen; }
            else { btnUplst.BackColor = Color.Yellow; }
            
           
            
            
        }
        private void timerFlush_Tick(object sender, EventArgs e)
        {
            if (BCKFlush.IsBusy == false)
            {
                object[] parameters = new object[] { towrite, gPathLogs, gpathECG, gpathECO, gpathHOLTER, gpathMAPA, gpathPE, gpathPDF, gDaystoFlush, gLastFlush , gpathUnconfirmed, gStoreReports , gStoreImages ,gStoreComplete };
                BCKFlush.RunWorkerAsync(parameters);
            }

        }
        private void timerFolders_Tick(object sender, EventArgs e)
        {
            if (BCKFolders.IsBusy == false)
            {
                btnFolders.BackColor = Color.LightGreen;
                object[] parameters = new object[] { towrite, gpathECG, gpathECO, gpathHOLTER, gpathMAPA, gpathPE, gECGRPrinter, gECGIPrinter, gECGCPrinter, gECORPrinter, gECOIPrinter, gECOCPrinter, gHOLTERRPrinter, gHOLTERIPrinter, gHOLTERCPrinter, gMAPARPrinter, gMAPAIPrinter, gMAPACPrinter, gprintECGR, gprintECGI, gprintECGC, gprintECOR, gprintECOI, gprintECOC, gprintHOLTERR, gprintHOLTERI, gprintHOLTERC, gprintMAPAR, gprintMAPAI, gprintMAPAC, gprintPER, gprintPEI, gprintPEC, gjoinECG, gjoinECO, gjoinHOLTER, gjoinMAPA, gjoinPE, gBookletPrinter, gBookletAudit, gBookletProcess, gBookletSecsToWait, gLastBookletprint, gPERPrinter, gPEIPrinter, gPECPrinter, gpathSumatraPDF , gStoreReports , gStoreImages ,gStoreComplete, gpathUnconfirmed, gdebug };
                BCKFolders.RunWorkerAsync(parameters);
            }
            else
            { btnFolders.BackColor = Color.Yellow; }
            if (BCKPDFOutput.IsBusy == false)
            {
                btnPDF.BackColor = Color.LightGreen;
                object[] parameters = new object[] { towrite, gpathPDF, gpathECG, gpathECGTbl, gpathECO, gpathECOTbl, gpathHOLTER, gpathHOLTERTbl, gpathMAPA, gpathMAPATbl, gpathPE, gpathPETbl, gpathUnconfirmed };
                BCKPDFOutput.RunWorkerAsync(parameters);
            }
            else
            { btnPDF.BackColor = Color.Yellow; }
        }
        private void timerLogger_Tick(object sender, EventArgs e)
        {
            if (BCKMasterWriter.IsBusy == true)
            { }
            else
            {

                object[] parameters = new object[] { towrite, gPathLogs };
                BCKMasterWriter.RunWorkerAsync(parameters);

            }
        }
        private void timerUpdateLst_Tick(object sender, EventArgs e)
        {
            if (BCKUpdateLst.IsBusy == false)
            {
                btnUplst.BackColor = Color.LightGreen;
                List<TreeNode> oldECG = new List<TreeNode>(GetAllNodes (this.treeViewECG));
                List<TreeNode> oldECO = new List<TreeNode>(GetAllNodes(this.treeViewECO));
                List<TreeNode> oldHOLTER = new List<TreeNode>(GetAllNodes(this.treeViewHOLTER));
                List<TreeNode> oldMAPA = new List<TreeNode>(GetAllNodes(this.treeViewMAPA));
                List<TreeNode> oldPE = new List<TreeNode>(GetAllNodes(this.treeViewPE));
                List<TreeNode> oldUC = new List<TreeNode>(GetAllNodes(this.treeViewUC));
                
                object[] parameters = new object[] { towrite, gpathECG, gpathECO, gpathHOLTER, gpathMAPA, gpathPE, gpathUnconfirmed, oldECG, oldECO, oldHOLTER, oldMAPA, oldPE, oldUC ,Gfirstload, gstate};
                BCKUpdateLst.RunWorkerAsync(parameters);
            }
        }
        private void timerEmail_Tick(object sender, EventArgs e)
        {
            if (BCKEmail.IsBusy == true)
            { btnEmail.BackColor = Color.Yellow; }
            else
            {
                btnEmail.BackColor = Color.LightGreen;
                object[] parameters = new object[] { towrite, gpathPDF, gPdfPrinter, gEmailInc, gpathHOLTERTAdd, gpathHOLTERTPass, gHOLTERExternalApp, gHOLTERExternalAppPath, gHOLTERSender, gHOLTERReceiver };
                BCKEmail.RunWorkerAsync(parameters);

            }
            if (BckReport.IsBusy == true)
            {  }
            else
            {
                
                object[] parameters = new object[] { towrite, gpathHOLTERTAdd, gpathHOLTERTPass, gHOLTERSender, gStoreComplete };
                BckReport.RunWorkerAsync(parameters);

            }
        }
        //BCK's
        private void BCKINIReader_DoWork(object sender, DoWorkEventArgs e)
        {
            Thread.Sleep(500);
            List <string> erros = new List<string>();
            string INIfilepathinterno = @"C:\PriJobAide\PriJobAide.ini";
            DateTime lastinicheckinterno = DateTime.MinValue;
            Helper internalhelp = new Helper();
            bool internaldebug = true;
            int internalstate = 2; // dois é o safestate onde só se visualiza alterações às pastas e só se altera aquilo que o utilizador pediu
            int internalDaystoFlush = 10;
            string internalpathPDF = "";
            string internalLocal = "";
            string internalpathSumatraPDF = "";
            string internalBookletPrinter = "";
            string internalBookletAudit = "";
            int internalBookletSecsToWait = 30;
            string internalBookletProcess = "";
            string internalpathECG = "";
            string internalECGRPrinter = "";
            bool internalgprintECGR = false;
            string internalECGIPrinter = "";
            bool internalgprintECGI = false;
            string internalECGCPrinter = "";
            bool internalgprintECGC = false;
            bool internalECGJoin = false;
            string internalpathECGTable = "";
            string internalpathECO = "";
            string internalECORPrinter = "";
            bool internalgprintECOR = false;
            string internalECOIPrinter = "";
            bool internalgprintECOI = false;
            string internalECOCPrinter = "";
            bool internalgprintECOC = false;
            bool internalECOJoin = false;
            string internalpathECOTable = "";
            string internalpathHOLTER = "";
            string internalHOLTERRPrinter = "";
            bool internalgprintHOLTERR = false;
            string internalHOLTERIPrinter = "";
            bool internalgprintHOLTERI = false;
            string internalHOLTERCPrinter = "";
            bool internalgprintHOLTERC = false;
            bool internalHOLTERJoin = false;
            string internalpathHOLTERTable = "";
            string internalHOLTEREmailAdd = "";
            string internalHOLTEREmailPass = "";
            DateTime internalHOLTEREmailStartCheck = DateTime.MinValue;
            DateTime internalHOLTEREmailEndCheck = DateTime.MinValue;
            string internalEmailInc = "";
            string internalpathMAPA = "";
            string internalMAPARPrinter = "";
            bool internalgprintMAPAR = false;
            string internalMAPAIPrinter = "";
            bool internalgprintMAPAI = false;
            string internalMAPACPrinter = "";
            bool internalgprintMAPAC = false;
            bool internalMAPAJoin = false;
            string internalpathMAPATable = "";
            string internalpathPE = "";
            string internalPERPrinter = "";
            bool internalgprintPER = false;
            string internalPEIPrinter = "";
            bool internalgprintPEI = false;
            string internalPECPrinter = "";
            bool internalgprintPEC = false;
            bool internalPEJoin = false;
            string internalpathPETable = "";
            string internalpathUnconfirmed = "";
            string internalStoreReports = "";
            string internalStoreImages = "";
            string internalStoreComplete = "";
            string internalPdfPrinter = "";
            string internalHOLTERSender = "";
            string internalHOLTERApp = "";
            string internalHOLTERAppPath = "";
            string internalEmailReceiver = "";
            object[] parameters = e.Argument as object[];
            ListLogger internallogger = (ListLogger)parameters[2];
            try
            {
                 INIfilepathinterno = (string)parameters[0];
                 lastinicheckinterno = (DateTime)parameters[1];
                 
                 
            }
            catch
            {
                erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao ler o ficheiro INI ao atribuir parametros."));
                // definir erro

            }
            FileInfo a = new FileInfo(INIfilepathinterno);
            if (File.Exists(INIfilepathinterno) == true)
            {
                

            }
            else
            { erros.Add (internalhelp.Formatmsg ("Ficheiro INI não existe."));}
                if (File.Exists(INIfilepathinterno) == true && internalhelp.FileInUse(INIfilepathinterno) == false && a.LastWriteTime != lastinicheckinterno)
            {
                int i = 0;
                
                erros.Add(internalhelp.Formatmsg("A carregar valores do ficheiro INI."));
                try
                {
                    List<string> inivalues = new List<string>();
                    inivalues = File.ReadAllLines(PathIni).ToList();
                    foreach (string s in inivalues)                
                    {
                        i++;
                        #region Debug
                        if (s.Contains("debug="))
                        {
                            try
                            {
                                internaldebug = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internaldebug = true; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'debug'.")); }
                            BCKINIReader.ReportProgress(1, internaldebug);
                            erros.Add(internalhelp.Formatmsg("debug = " + internaldebug.ToString()));
                        } 
                        #endregion
                        #region State
                        if (s.Contains("state="))
                        {
                            try
                            {
                                internalstate = Convert.ToInt32(internalhelp.TrimIniText(s));
                            }
                            catch { internalstate = 2; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'state'.")); }
                            BCKINIReader.ReportProgress(2, internalstate);
                            erros.Add(internalhelp.Formatmsg("state = " + internalstate.ToString()));
                        } 
                        #endregion
                        #region DaystoFlush
                        if (s.Contains("DaystoFlush="))
                        {
                            try
                            {
                                internalDaystoFlush = Convert.ToInt32(internalhelp.TrimIniText(s));
                            }
                            catch { internalDaystoFlush = 10; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'DaystoFlush'.")); }
                            BCKINIReader.ReportProgress(3, internalDaystoFlush);
                            erros.Add(internalhelp.Formatmsg("DaystoFlush = " + internalDaystoFlush.ToString()));
                        } 
                        #endregion
                        #region pathPDF
                        if (s.Contains("pathPDF="))
                        {
                            try
                            {
                                internalpathPDF = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathPDF = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathPDF'.")); }
                            BCKINIReader.ReportProgress(4, internalpathPDF);
                            erros.Add(internalhelp.Formatmsg("pathPDF = " + internalpathPDF));
                        } 
                        #endregion 
                        #region pathSumatraPDF
                        if (s.Contains("pathSumatraPDF="))
                        {
                            try
                            {
                                internalpathSumatraPDF = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathSumatraPDF = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathSumatraPDF'.")); }
                            BCKINIReader.ReportProgress(5, internalpathSumatraPDF);
                            erros.Add(internalhelp.Formatmsg("pathSumatraPDF = " + internalpathSumatraPDF));
                        }
                        #endregion
                        #region BookletPrinter
                        if (s.Contains("BookletPrinter="))
                        {
                            try
                            {
                                internalBookletPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalBookletPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'BookletPrinter'.")); }
                            BCKINIReader.ReportProgress(6, internalBookletPrinter);
                            erros.Add(internalhelp.Formatmsg("BookletPrinter = " + internalBookletPrinter));
                        }
                        #endregion
                        #region BookletAudit
                        if (s.Contains("BookletAudit="))
                        {
                            try
                            {
                                internalBookletAudit = internalhelp.TrimIniText(s);
                            }
                            catch { internalBookletAudit = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'BookletAudit'.")); }
                            BCKINIReader.ReportProgress(7, internalBookletAudit);
                            erros.Add(internalhelp.Formatmsg("BookletAudit = " + internalBookletAudit));
                        }
                        #endregion
                        #region BookletSecsToWait
                        if (s.Contains("BookletSecsToWait="))
                        {
                            try
                            {
                                internalBookletSecsToWait = Convert.ToInt32 (internalhelp.TrimIniText(s));
                            }
                            catch { internalBookletSecsToWait = 30; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'BookletSecsToWait'.")); }
                            BCKINIReader.ReportProgress(8, internalBookletSecsToWait);
                            erros.Add(internalhelp.Formatmsg("BookletSecsToWait = " + internalBookletSecsToWait.ToString()));
                        }
                        #endregion
                        #region BookletProcess
                        if (s.Contains("BookletProcess="))
                        {
                            try
                            {
                                internalBookletProcess = internalhelp.TrimIniText(s);
                            }
                            catch { internalBookletProcess = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'BookletProcess'.")); }
                            BCKINIReader.ReportProgress(9, internalBookletProcess);
                            erros.Add(internalhelp.Formatmsg("BookletProcess = " + internalBookletProcess));
                        }
                        #endregion
                        #region pathECG
                        if (s.Contains("pathECG="))
                        {
                            try
                            {
                                internalpathECG = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathECG = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathECG'.")); }
                            BCKINIReader.ReportProgress(10, internalpathECG);
                            erros.Add(internalhelp.Formatmsg("pathECG = " + internalpathECG));
                            
                        }  
                        #endregion

                        #region ECGPrintR
                        if (s.Contains("ECGPrintR="))
                        {
                            try
                            {
                                internalgprintECGR = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintECGR = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECGPrintR'.")); }
                            BCKINIReader.ReportProgress(11, internalgprintECGR);
                            erros.Add(internalhelp.Formatmsg("ECGPrintR = " + internalgprintECGR));
                        }   
                        #endregion
                        #region ECGRPrinter
                        if (s.Contains("ECGRPrinter="))
                        {
                            try
                            {
                                internalECGRPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalECGRPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECGRPrinter'.")); }
                            BCKINIReader.ReportProgress(12, internalECGRPrinter);
                            erros.Add(internalhelp.Formatmsg("ECGRPrinter = " + internalECGRPrinter));
                        }    
                        #endregion
                        #region ECGPrintI
                        if (s.Contains("ECGPrintI="))
                        {
                            try
                            {
                                internalgprintECGI = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintECGI = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECGPrintI'.")); }
                            BCKINIReader.ReportProgress(13, internalgprintECGI);
                            erros.Add(internalhelp.Formatmsg("ECGPrintI = " + internalgprintECGI));
                        }
                        #endregion
                        #region ECGIPrinter
                        if (s.Contains("ECGIPrinter="))
                        {
                            try
                            {
                                internalECGIPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalECGIPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECGIPrinter'.")); }
                            BCKINIReader.ReportProgress(14, internalECGIPrinter);
                            erros.Add(internalhelp.Formatmsg("ECGIPrinter = " + internalECGIPrinter));
                        }
                        #endregion
                        #region ECGPrintC
                        if (s.Contains("ECGPrintC="))
                        {
                            try
                            {
                                internalgprintECGC = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintECGC = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECGPrintC'.")); }
                            BCKINIReader.ReportProgress(15, internalgprintECGC);
                            erros.Add(internalhelp.Formatmsg("ECGPrintC = " + internalgprintECGC));
                        }
                        #endregion
                        #region ECGCPrinter
                        if (s.Contains("ECGCPrinter="))
                        {
                            try
                            {
                                internalECGCPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalECGCPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECGCPrinter'.")); }
                            BCKINIReader.ReportProgress(16, internalECGCPrinter);
                            erros.Add(internalhelp.Formatmsg("ECGCPrinter = " + internalECGCPrinter));
                        }
                        #endregion
                        #region ECGJoin
                        if (s.Contains("ECGJoin="))
                        {
                            try
                            {
                                internalECGJoin = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalECGJoin = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECGJoin'.")); }
                            BCKINIReader.ReportProgress(17, internalECGJoin);
                            erros.Add(internalhelp.Formatmsg("ECGJoin = " + internalECGJoin));
                        }
                        #endregion
                        #region pathECGTable
                        if (s.Contains("pathECGTable="))
                        {
                            try
                            {
                                internalpathECGTable = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathECGTable = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathECGTable'.")); }
                            BCKINIReader.ReportProgress(18, internalpathECGTable);
                            erros.Add(internalhelp.Formatmsg("pathECGTable = " + internalpathECGTable));
                        }
                        #endregion
                        // here
                        #region pathECO
                        if (s.Contains("pathECO="))
                        {
                            try
                            {
                                internalpathECO = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathECO = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathECO'.")); }
                            BCKINIReader.ReportProgress(19, internalpathECO);
                            erros.Add(internalhelp.Formatmsg("pathECO = " + internalpathECO));
                            
                        }
                        #endregion
                        #region ECOPrintR
                        if (s.Contains("ECOPrintR="))
                        {
                            try
                            {
                                internalgprintECOR = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintECOR = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECOPrintR'.")); }
                            BCKINIReader.ReportProgress(20, internalgprintECOR);
                            erros.Add(internalhelp.Formatmsg("ECOPrintR = " + internalgprintECOR));
                        }
                        #endregion
                        #region ECORPrinter
                        if (s.Contains("ECORPrinter="))
                        {
                            try
                            {
                                internalECORPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalECORPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECORPrinter'.")); }
                            BCKINIReader.ReportProgress(21, internalECORPrinter);
                            erros.Add(internalhelp.Formatmsg("ECORPrinter = " + internalECORPrinter));
                        }
                        #endregion
                        #region ECOPrintI
                        if (s.Contains("ECOPrintI="))
                        {
                            try
                            {
                                internalgprintECOI = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintECOI = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECOPrintI'.")); }
                            BCKINIReader.ReportProgress(22, internalgprintECOI);
                            erros.Add(internalhelp.Formatmsg("ECOPrintI = " + internalgprintECOI));
                        }
                        #endregion
                        #region ECOIPrinter
                        if (s.Contains("ECOIPrinter="))
                        {
                            try
                            {
                                internalECOIPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalECOIPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECOIPrinter'.")); }
                            BCKINIReader.ReportProgress(23, internalECOIPrinter);
                            erros.Add(internalhelp.Formatmsg("ECOIPrinter = " + internalECOIPrinter));
                        }
                        #endregion
                        #region ECOPrintC
                        if (s.Contains("ECOPrintC="))
                        {
                            try
                            {
                                internalgprintECOC = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintECOC = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECOPrintC'.")); }
                            BCKINIReader.ReportProgress(24, internalgprintECOC);
                            erros.Add(internalhelp.Formatmsg("ECOPrintC = " + internalgprintECOC));
                        }
                        #endregion
                        #region ECOCPrinter
                        if (s.Contains("ECOCPrinter="))
                        {
                            try
                            {
                                internalECOCPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalECOCPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECOCPrinter'.")); }
                            BCKINIReader.ReportProgress(25, internalECOCPrinter);
                            erros.Add(internalhelp.Formatmsg("ECOCPrinter = " + internalECOCPrinter));
                        }
                        #endregion
                        #region ECOJoin
                        if (s.Contains("ECOJoin="))
                        {
                            try
                            {
                                internalECOJoin = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalECOJoin = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'ECOJoin'.")); }
                            BCKINIReader.ReportProgress(26, internalECOJoin);
                            erros.Add(internalhelp.Formatmsg("ECOJoin = " + internalECOJoin));
                        }
                        #endregion
                        #region pathECOTable
                        if (s.Contains("pathECOTable="))
                        {
                            try
                            {
                                internalpathECOTable = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathECOTable = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathECOTable'.")); }
                            BCKINIReader.ReportProgress(27, internalpathECOTable);
                            erros.Add(internalhelp.Formatmsg("pathECOTable = " + internalpathECOTable));
                        }
                        #endregion
                        //here 2
                        #region pathHOLTER
                        if (s.Contains("pathHOLTER="))
                        {
                            try
                            {
                                internalpathHOLTER = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathHOLTER = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathHOLTER'.")); }
                            BCKINIReader.ReportProgress(28, internalpathHOLTER);
                            erros.Add(internalhelp.Formatmsg("pathHOLTER = " + internalpathHOLTER));
                            
                        }
                        #endregion
                        #region HOLTERPrintR
                        if (s.Contains("HOLTERPrintR="))
                        {
                            try
                            {
                                internalgprintHOLTERR = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintHOLTERR = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'HOLTERPrintR'.")); }
                            BCKINIReader.ReportProgress(29, internalgprintHOLTERR);
                            erros.Add(internalhelp.Formatmsg("HOLTERPrintR = " + internalgprintHOLTERR));
                        }
                        #endregion
                        #region HOLTERRPrinter
                        if (s.Contains("HOLTERRPrinter="))
                        {
                            try
                            {
                                internalHOLTERRPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalHOLTERRPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'HOLTERRPrinter'.")); }
                            BCKINIReader.ReportProgress(30, internalHOLTERRPrinter);
                            erros.Add(internalhelp.Formatmsg("HOLTERRPrinter = " + internalHOLTERRPrinter));
                        }
                        #endregion
                        #region HOLTERPrintI
                        if (s.Contains("HOLTERPrintI="))
                        {
                            try
                            {
                                internalgprintHOLTERI = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintHOLTERI = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'HOLTERPrintI'.")); }
                            BCKINIReader.ReportProgress(31, internalgprintHOLTERI);
                            erros.Add(internalhelp.Formatmsg("HOLTERPrintI = " + internalgprintHOLTERI));
                        }
                        #endregion
                        #region HOLTERIPrinter
                        if (s.Contains("HOLTERIPrinter="))
                        {
                            try
                            {
                                internalHOLTERIPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalHOLTERIPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'HOLTERIPrinter'.")); }
                            BCKINIReader.ReportProgress(32, internalHOLTERIPrinter);
                            erros.Add(internalhelp.Formatmsg("HOLTERIPrinter = " + internalHOLTERIPrinter));
                        }
                        #endregion
                        #region HOLTERPrintC
                        if (s.Contains("HOLTERPrintC="))
                        {
                            try
                            {
                                internalgprintHOLTERC = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintHOLTERC = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'HOLTERPrintC'.")); }
                            BCKINIReader.ReportProgress(33, internalgprintHOLTERC);
                            erros.Add(internalhelp.Formatmsg("HOLTERPrintC = " + internalgprintHOLTERC));
                        }
                        #endregion
                        #region HOLTERCPrinter
                        if (s.Contains("HOLTERCPrinter="))
                        {
                            try
                            {
                                internalHOLTERCPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalHOLTERCPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'HOLTERCPrinter'.")); }
                            BCKINIReader.ReportProgress(34, internalHOLTERCPrinter);
                            erros.Add(internalhelp.Formatmsg("HOLTERCPrinter = " + internalHOLTERCPrinter));
                        }
                        #endregion
                        #region HOLTERJoin
                        if (s.Contains("HOLTERJoin="))
                        {
                            try
                            {
                                internalHOLTERJoin = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalHOLTERJoin = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'HOLTERJoin'.")); }
                            BCKINIReader.ReportProgress(35, internalHOLTERJoin);
                            erros.Add(internalhelp.Formatmsg("HOLTERJoin = " + internalHOLTERJoin));
                        }
                        #endregion
                        #region pathHOLTERTable
                        if (s.Contains("pathHOLTERTable="))
                        {
                            try
                            {
                                internalpathHOLTERTable = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathHOLTERTable = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathHOLTERTable'.")); }
                            BCKINIReader.ReportProgress(36, internalpathHOLTERTable);
                            erros.Add(internalhelp.Formatmsg("pathHOLTERTable = " + internalpathHOLTERTable));
                        }
                        #endregion
                        #region HOLTEREmailAdd
                        if (s.Contains("HOLTEREmailAdd="))
                        {
                            try
                            {
                                internalHOLTEREmailAdd = internalhelp.TrimIniText(s);
                                
                            }
                            catch { internalpathHOLTERTable = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'HOLTEREmailAdd'.")); }
                            BCKINIReader.ReportProgress(37, internalHOLTEREmailAdd);
                            erros.Add(internalhelp.Formatmsg("HOLTEREmailAdd = " + internalHOLTEREmailAdd));
                        }
                        #endregion
                        #region HOLTEREmailPass
                        if (s.Contains("HOLTEREmailPass="))
                        {
                            try
                            {
                                internalHOLTEREmailPass = internalhelp.TrimIniText(s);

                            }
                            catch { internalHOLTEREmailPass = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'HOLTEREmailPass'.")); }
                            BCKINIReader.ReportProgress(38, internalHOLTEREmailPass);
                            erros.Add(internalhelp.Formatmsg("HOLTEREmailPass = " + internalHOLTEREmailPass));
                        }
                        #endregion
                        
                        
                        // here 3
                        #region pathMAPA
                        if (s.Contains("pathMAPA="))
                        {
                            try
                            {
                                internalpathMAPA = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathMAPA = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathMAPA'.")); }
                            BCKINIReader.ReportProgress(41, internalpathMAPA);
                            erros.Add(internalhelp.Formatmsg("pathMAPA = " + internalpathMAPA));
                           
                        }
                        #endregion
                        #region MAPAPrintR
                        if (s.Contains("MAPAPrintR="))
                        {
                            try
                            {
                                internalgprintMAPAR = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintMAPAR = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'MAPAPrintR'.")); }
                            BCKINIReader.ReportProgress(42, internalgprintMAPAR);
                            erros.Add(internalhelp.Formatmsg("MAPAPrintR = " + internalgprintMAPAR));
                        }
                        #endregion
                        #region MAPARPrinter
                        if (s.Contains("MAPARPrinter="))
                        {
                            try
                            {
                                internalMAPARPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalMAPARPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'MAPARPrinter'.")); }
                            BCKINIReader.ReportProgress(43, internalMAPARPrinter);
                            erros.Add(internalhelp.Formatmsg("MAPARPrinter = " + internalMAPARPrinter));
                        }
                        #endregion
                        #region MAPAPrintI
                        if (s.Contains("MAPAPrintI="))
                        {
                            try
                            {
                                internalgprintMAPAI = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintMAPAI = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'MAPAPrintI'.")); }
                            BCKINIReader.ReportProgress(44, internalgprintMAPAI);
                            erros.Add(internalhelp.Formatmsg("MAPAPrintI = " + internalgprintMAPAI));
                        }
                        #endregion
                        #region MAPAIPrinter
                        if (s.Contains("MAPAIPrinter="))
                        {
                            try
                            {
                                internalMAPAIPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalMAPAIPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'MAPAIPrinter'.")); }
                            BCKINIReader.ReportProgress(45, internalMAPAIPrinter);
                            erros.Add(internalhelp.Formatmsg("MAPAIPrinter = " + internalMAPAIPrinter));
                        }
                        #endregion
                        #region MAPAPrintC
                        if (s.Contains("MAPAPrintC="))
                        {
                            try
                            {
                                internalgprintMAPAC = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintMAPAC = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'MAPAPrintC'.")); }
                            BCKINIReader.ReportProgress(46, internalgprintMAPAC);
                            erros.Add(internalhelp.Formatmsg("MAPAPrintC = " + internalgprintMAPAC));
                        }
                        #endregion
                        #region MAPACPrinter
                        if (s.Contains("MAPACPrinter="))
                        {
                            try
                            {
                                internalMAPACPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalMAPACPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'MAPACPrinter'.")); }
                            BCKINIReader.ReportProgress(47, internalMAPACPrinter);
                            erros.Add(internalhelp.Formatmsg("MAPACPrinter = " + internalMAPACPrinter));
                        }
                        #endregion
                        #region MAPAJoin
                        if (s.Contains("MAPAJoin="))
                        {
                            try
                            {
                                internalMAPAJoin = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalMAPAJoin = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'MAPAJoin'.")); }
                            BCKINIReader.ReportProgress(48, internalMAPAJoin);
                            erros.Add(internalhelp.Formatmsg("MAPAJoin = " + internalMAPAJoin));
                        }
                        #endregion
                        #region pathMAPATable
                        if (s.Contains("pathMAPATable="))
                        {
                            try
                            {
                                internalpathMAPATable = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathMAPATable = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathMAPATable'.")); }
                            BCKINIReader.ReportProgress(49, internalpathMAPATable);
                            erros.Add(internalhelp.Formatmsg("pathMAPATable = " + internalpathMAPATable));
                        }
                        #endregion
                        // HERE
                        #region pathPE
                        if (s.Contains("pathPE="))
                        {
                            try
                            {
                                internalpathPE = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathPE = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathPE'.")); }
                            BCKINIReader.ReportProgress(50, internalpathPE);
                            erros.Add(internalhelp.Formatmsg("pathPE = " + internalpathPE));
                            
                        }
                        #endregion
                        #region PEPrintR
                        if (s.Contains("PEPrintR="))
                        {
                            try
                            {
                                internalgprintPER = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintPER = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'PEPrintR'.")); }
                            BCKINIReader.ReportProgress(51, internalgprintPER);
                            erros.Add(internalhelp.Formatmsg("PEPrintR = " + internalgprintPER));
                        }
                        #endregion
                        #region PERPrinter
                        if (s.Contains("PERPrinter="))
                        {
                            try
                            {
                                internalPERPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalPERPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'PERPrinter'.")); }
                            BCKINIReader.ReportProgress(52, internalPERPrinter);
                            erros.Add(internalhelp.Formatmsg("PERPrinter = " + internalPERPrinter));
                        }
                        #endregion
                        #region PEPrintI
                        if (s.Contains("PEPrintI="))
                        {
                            try
                            {
                                internalgprintPEI = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintPEI = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'PEPrintI'.")); }
                            BCKINIReader.ReportProgress(53, internalgprintPEI);
                            erros.Add(internalhelp.Formatmsg("PEPrintI = " + internalgprintPEI));
                        }
                        #endregion
                        #region PEIPrinter
                        if (s.Contains("PEIPrinter="))
                        {
                            try
                            {
                                internalPEIPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalPEIPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'PEIPrinter'.")); }
                            BCKINIReader.ReportProgress(54, internalPEIPrinter);
                            erros.Add(internalhelp.Formatmsg("PEIPrinter = " + internalPEIPrinter));
                        }
                        #endregion
                        #region PEPrintC
                        if (s.Contains("PEPrintC="))
                        {
                            try
                            {
                                internalgprintPEC = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalgprintPEC = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'PEPrintC'.")); }
                            BCKINIReader.ReportProgress(55, internalgprintPEC);
                            erros.Add(internalhelp.Formatmsg("PEPrintC = " + internalgprintPEC));
                        }
                        #endregion
                        #region PECPrinter
                        if (s.Contains("PECPrinter="))
                        {
                            try
                            {
                                internalPECPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalPECPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'PECPrinter'.")); }
                            BCKINIReader.ReportProgress(56, internalPECPrinter);
                            erros.Add(internalhelp.Formatmsg("PECPrinter = " + internalPECPrinter));
                        }
                        #endregion
                        #region PEJoin
                        if (s.Contains("PEJoin="))
                        {
                            try
                            {
                                internalPEJoin = Convert.ToBoolean(internalhelp.TrimIniText(s));
                            }
                            catch { internalPEJoin = false; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'PEJoin'.")); }
                            BCKINIReader.ReportProgress(57, internalPEJoin);
                            erros.Add(internalhelp.Formatmsg("PEJoin = " + internalPEJoin));
                        }
                        #endregion
                        #region pathPETable
                        if (s.Contains("pathPETable="))
                        {
                            try
                            {
                                internalpathPETable = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathPETable = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathPETable'.")); }
                            BCKINIReader.ReportProgress(58, internalpathPETable);
                            erros.Add(internalhelp.Formatmsg("pathPETable = " + internalpathPETable));
                        }
                        #endregion
                        if (s.Contains("HOLTEREmailReceiver="))
                        {
                            try
                            {
                                internalEmailReceiver = internalhelp.TrimIniText(s);
                            }
                            catch { internalEmailReceiver = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'internalEmailReceiver'.")); }
                            BCKINIReader.ReportProgress(68, internalEmailReceiver);
                            erros.Add(internalhelp.Formatmsg("HOLTEREmailReceiver = " + internalEmailReceiver));
                        }
                        if (s.Contains("HOLTEREmailSender="))
                        {
                            try
                            {
                                internalHOLTERSender = internalhelp.TrimIniText(s);
                            }
                            catch { internalHOLTERSender = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'internalHOLTERSender'.")); }
                            BCKINIReader.ReportProgress(67, internalHOLTERSender);
                            erros.Add(internalhelp.Formatmsg("HOLTEREmailSender = " + internalHOLTERSender));
                        }
                        if (s.Contains("HolterExternalAppPath="))
                        {
                            try
                            {
                                internalHOLTERAppPath = internalhelp.TrimIniText(s);
                            }
                            catch { internalHOLTERAppPath = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'internalHOLTERAppPath'.")); }
                            BCKINIReader.ReportProgress(66, internalHOLTERAppPath);
                            erros.Add(internalhelp.Formatmsg("HolterExternalAppPath = " + internalHOLTERAppPath));
                        }
                        if (s.Contains("HolterExternalAppProcess="))
                        {
                            try
                            {
                                internalHOLTERApp = internalhelp.TrimIniText(s);
                            }
                            catch { internalHOLTERApp = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'internalHOLTERApp'.")); }
                            BCKINIReader.ReportProgress(65, internalHOLTERApp);
                            erros.Add(internalhelp.Formatmsg("HolterExternalAppProcess = " + internalHOLTERApp));
                        }
                        if (s.Contains("HolterEmailPath="))
                        {
                            try
                            {
                                internalEmailInc = internalhelp.TrimIniText(s);
                            }
                            catch { internalEmailInc = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'internalEmailInc'.")); }
                            BCKINIReader.ReportProgress(63, internalEmailInc);
                            erros.Add(internalhelp.Formatmsg("HolterEmailPath = " + internalEmailInc));
                        }
                        if (s.Contains("PdfPrinter="))
                        {
                            try
                            {
                                internalPdfPrinter = internalhelp.TrimIniText(s);
                            }
                            catch { internalPdfPrinter = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'internalPdfPrinter'.")); }
                            BCKINIReader.ReportProgress(64, internalPdfPrinter);
                            erros.Add(internalhelp.Formatmsg("PdfPrinter = " + internalPdfPrinter));
                        }
                        if (s.Contains("pathUnconfirmed="))
                        {
                            try
                            {
                                internalpathUnconfirmed = internalhelp.TrimIniText(s);
                            }
                            catch { internalpathUnconfirmed = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'pathUnconfirmed'.")); }
                            BCKINIReader.ReportProgress(59, internalpathUnconfirmed);
                            erros.Add(internalhelp.Formatmsg("pathUnconfirmed = " + internalpathUnconfirmed));
                            
                        }
                        if (s.Contains("PathStoreReports="))
                        {
                            try
                            {
                                internalStoreReports = internalhelp.TrimIniText(s);
                            }
                            catch { internalStoreReports = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'StoreReports'.")); }
                            BCKINIReader.ReportProgress(60, internalStoreReports);
                            erros.Add(internalhelp.Formatmsg("PathStoreReports = " + internalStoreReports));
                        }
                        if (s.Contains("PathStoreImages="))
                        {
                            try
                            {
                                internalStoreImages = internalhelp.TrimIniText(s);
                            }
                            catch { internalStoreImages = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'StoreImages'.")); }
                            BCKINIReader.ReportProgress(61, internalStoreImages);
                            erros.Add(internalhelp.Formatmsg("PathStoreImages = " + internalStoreImages));
                        }
                        if (s.Contains("PathStoreComplete="))
                        {
                            try
                            {
                                internalStoreComplete = internalhelp.TrimIniText(s);
                            }
                            catch { internalStoreComplete = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'StoreComplete'.")); }
                            BCKINIReader.ReportProgress(62, internalStoreComplete);
                            erros.Add(internalhelp.Formatmsg("StoreComplete = " + internalStoreComplete));
                        }
                        if (s.Contains("local="))
                        {
                            try
                            {
                                internalLocal = internalhelp.TrimIniText(s);
                            }
                            catch { internalLocal = ""; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao converter o parametro 'internalLocal'.")); }
                            BCKINIReader.ReportProgress(69, internalLocal);
                            erros.Add(internalhelp.Formatmsg("internalLocal = " + internalLocal));
                        }
                    }
                    
                }
                catch
                {
                    erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao ler o ficheiro INI."));
                    // definir erro
                }
                
                if (i == 0)
                {
                    BCKINIReader.ReportProgress(0);
                    erros.Add(internalhelp.Formatmsg("Ficheiro INI está vazio."));
                    Thread.Sleep(1000);
                }
                lastinicheckinterno = a.LastWriteTime;
                // falta lista de erros
                 object[] resultados = new object[] { lastinicheckinterno };
                 if (erros.Count > 0)
                 {
                     internallogger.AddEx(erros);
                 }
                 
                e.Result = resultados;
            }
            
        }
        private void BCKINIReader_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            #region Percentage = 0
            if (e.ProgressPercentage == 0)
            {
                btnIni.BackColor = Color.DarkRed;
                try
                {
                    frmsetup.Show();
                }
                catch
                {
                    frmsetup = new FrmSetup();
                    frmsetup.Show();
                }
            } 
            #endregion
            #region Progress Report
            if (e.ProgressPercentage == 68)
            {
                object Userstate = e.UserState as object;
                string internalEmailReceiver = (string)e.UserState;
                gHOLTERReceiver = internalEmailReceiver;
                gimediate = MyHelper.Formatmsg("HOLTEREmailReceiver=" + internalEmailReceiver);
            }
            if (e.ProgressPercentage == 69)
            {
                object Userstate = e.UserState as object;
                string internalLocal = (string)e.UserState;
                gLocal = internalLocal;
                gimediate = MyHelper.Formatmsg("Local=" + internalLocal);
            }
            if (e.ProgressPercentage == 65)
            {
                object Userstate = e.UserState as object;
                string internalHOLTERApp = (string)e.UserState;
                gHOLTERExternalApp = internalHOLTERApp;
                gimediate = MyHelper.Formatmsg("HOLTERExternalApp=" + internalHOLTERApp);
            }
            if (e.ProgressPercentage == 66)
            {
                object Userstate = e.UserState as object;
                string internalHOLTERAppPath = (string)e.UserState;
                gHOLTERExternalAppPath = internalHOLTERAppPath;
                gimediate = MyHelper.Formatmsg("HOLTERExternalAppPath=" + internalHOLTERAppPath);
            }
            if (e.ProgressPercentage == 67)
            {
                object Userstate = e.UserState as object;
                string internalHOLTERSender = (string)e.UserState;
                gHOLTERSender = internalHOLTERSender;
                gimediate = MyHelper.Formatmsg("HOLTERSender=" + internalHOLTERSender);
            }
            if (e.ProgressPercentage == 1)
            {
                object Userstate = e.UserState as object;
                bool InternalDebug = (bool)e.UserState;
                gdebug = InternalDebug;
                gimediate = MyHelper.Formatmsg("debug=" + InternalDebug.ToString());
            }
            if (e.ProgressPercentage == 2)
            {
                object Userstate = e.UserState as object;
                int internalstate = (int)e.UserState;
                if (gstate != 2)
                {
                    gstate = internalstate;
                } gimediate = MyHelper.Formatmsg("state=" + internalstate.ToString());
            }
            if (e.ProgressPercentage == 3)
            {
                object Userstate = e.UserState as object;
                int internalDaystoFlush = (int)e.UserState;
                gDaystoFlush = internalDaystoFlush;
                gimediate = MyHelper.Formatmsg("DaystoFlush=" + internalDaystoFlush.ToString());
            }
            if (e.ProgressPercentage == 4)
            {
                object Userstate = e.UserState as object;
                string internalpathPDF = (string)e.UserState;
                gpathPDF = internalpathPDF;
                gimediate = MyHelper.Formatmsg("pathPDF=" + internalpathPDF);
            }
            if (e.ProgressPercentage == 5)
            {
                object Userstate = e.UserState as object;
                string internalpathSumatraPDF = (string)e.UserState;
                gpathSumatraPDF = internalpathSumatraPDF;
                gimediate = MyHelper.Formatmsg("pathSumatraPDF=" + internalpathSumatraPDF);
            }
            if (e.ProgressPercentage == 6)
            {
                object Userstate = e.UserState as object;
                string internalBookletPrinter = (string)e.UserState;
                gBookletPrinter = internalBookletPrinter;
                gimediate = MyHelper.Formatmsg("BookletPrinter=" + internalBookletPrinter);
            }
            if (e.ProgressPercentage == 7)
            {
                object Userstate = e.UserState as object;
                string internalBookletAudit = (string)e.UserState;
                this.gBookletAudit = internalBookletAudit;
                gimediate = MyHelper.Formatmsg("BookletAudit=" + internalBookletAudit);
            }
            if (e.ProgressPercentage == 8)
            {
                object Userstate = e.UserState as object;
                int internalBookletSecsToWait = (int)e.UserState;
                gBookletSecsToWait = internalBookletSecsToWait;
                gimediate = MyHelper.Formatmsg("BookletSecsToWait=" + internalBookletSecsToWait.ToString());
            }
            if (e.ProgressPercentage == 9)
            {
                object Userstate = e.UserState as object;
                string internalBookletProcess = (string)e.UserState;
                gBookletProcess = internalBookletProcess;
                gimediate = MyHelper.Formatmsg("BookletProcess=" + internalBookletProcess);
            }
            //ECG
            if (e.ProgressPercentage == 10)
            {
                object Userstate = e.UserState as object;
                string internalPathECG = (string)e.UserState;
                gpathECG = internalPathECG;
                gimediate = MyHelper.Formatmsg("pathECG=" + internalPathECG);
            }
            if (e.ProgressPercentage == 11)
            {
                object Userstate = e.UserState as object;
                bool internalprintECGR = (bool)e.UserState;
                this.gprintECGR = internalprintECGR;
                gimediate = MyHelper.Formatmsg("printECGR=" + internalprintECGR.ToString());
            }
            if (e.ProgressPercentage == 12)
            {
                object Userstate = e.UserState as object;
                string internalECGRPrinter = (string)e.UserState;
                this.gECGRPrinter = internalECGRPrinter;
                gimediate = MyHelper.Formatmsg("ECGRPrinter=" + internalECGRPrinter);
            }
            if (e.ProgressPercentage == 13)
            {
                object Userstate = e.UserState as object;
                bool internalprintECGI = (bool)e.UserState;
                this.gprintECGI = internalprintECGI;
                gimediate = MyHelper.Formatmsg("printECGI=" + internalprintECGI.ToString());
            }
            if (e.ProgressPercentage == 14)
            {
                object Userstate = e.UserState as object;
                string internalECGIPrinter = (string)e.UserState;
                this.gECGIPrinter = internalECGIPrinter;
                gimediate = MyHelper.Formatmsg("printECGI=" + internalECGIPrinter);
            }
            if (e.ProgressPercentage == 15)
            {
                object Userstate = e.UserState as object;
                bool internalprintECGC = (bool)e.UserState;
                this.gprintECGC = internalprintECGC;
                gimediate = MyHelper.Formatmsg("printECGC=" + internalprintECGC.ToString());
            }
            if (e.ProgressPercentage == 16)
            {
                object Userstate = e.UserState as object;
                string internalECGCPrinter = (string)e.UserState;
                this.gECGCPrinter = internalECGCPrinter;
                gimediate = MyHelper.Formatmsg("ECGCPrinter=" + internalECGCPrinter);
            }
            if (e.ProgressPercentage == 17)
            {
                object Userstate = e.UserState as object;
                bool internaljoinECG = (bool)e.UserState;
                this.gjoinECG = internaljoinECG;
                gimediate = MyHelper.Formatmsg("joinECG=" + internaljoinECG.ToString());
            }
            if (e.ProgressPercentage == 18)
            {
                object Userstate = e.UserState as object;
                string internalprintECGTbl = (string)e.UserState;
                this.gpathECGTbl = internalprintECGTbl;
                gimediate = MyHelper.Formatmsg("pathECGTbl=" + internalprintECGTbl);
            }
            //END ECG
            //START ECO
            if (e.ProgressPercentage == 19)
            {
                object Userstate = e.UserState as object;
                string internalPathECO = (string)e.UserState;
                gpathECO = internalPathECO;
                gimediate = MyHelper.Formatmsg("PathECO=" + internalPathECO);
            }
            if (e.ProgressPercentage == 20)
            {
                object Userstate = e.UserState as object;
                bool internalprintECOR = (bool)e.UserState;
                this.gprintECOR = internalprintECOR;
                gimediate = MyHelper.Formatmsg("printECOR=" + internalprintECOR.ToString());
            }
            if (e.ProgressPercentage == 21)
            {
                object Userstate = e.UserState as object;
                string internalprintECOR = (string)e.UserState;
                this.gECORPrinter = internalprintECOR;
                gimediate = MyHelper.Formatmsg("ECORPrinter=" + internalprintECOR.ToString());
            }
            if (e.ProgressPercentage == 22)
            {
                object Userstate = e.UserState as object;
                bool internalprintECOI = (bool)e.UserState;
                this.gprintECOI = internalprintECOI;
                gimediate = MyHelper.Formatmsg("printECOI=" + internalprintECOI.ToString());
            }
            if (e.ProgressPercentage == 23)
            {
                object Userstate = e.UserState as object;
                string internalprintECOI = (string)e.UserState;
                this.gECOIPrinter = internalprintECOI;
                gimediate = MyHelper.Formatmsg("ECOIPrinter=" + internalprintECOI.ToString());
            }
            if (e.ProgressPercentage == 24)
            {
                object Userstate = e.UserState as object;
                bool internalprintECOC = (bool)e.UserState;
                this.gprintECOC = internalprintECOC;
                gimediate = MyHelper.Formatmsg("printECOC=" + internalprintECOC.ToString());
            }
            if (e.ProgressPercentage == 25)
            {
                object Userstate = e.UserState as object;
                string internalprintECOC = (string)e.UserState;
                this.gECOCPrinter = internalprintECOC;
                gimediate = MyHelper.Formatmsg("ECOCPrinter=" + internalprintECOC.ToString());
            }
            if (e.ProgressPercentage == 26)
            {
                object Userstate = e.UserState as object;
                bool internaljoinECO = (bool)e.UserState;
                this.gjoinECO = internaljoinECO;
                gimediate = MyHelper.Formatmsg("joinECO=" + internaljoinECO.ToString());
            }
            if (e.ProgressPercentage == 27)
            {
                object Userstate = e.UserState as object;
                string internalprintECOTbl = (string)e.UserState;
                this.gpathECOTbl = internalprintECOTbl;
                gimediate = MyHelper.Formatmsg("pathECOTbl=" + internalprintECOTbl.ToString());
            }
            //END ECO
            //START HOLTER
            if (e.ProgressPercentage == 28)
            {
                object Userstate = e.UserState as object;
                string internalPathHOLTER = (string)e.UserState;
                gpathHOLTER = internalPathHOLTER;
                gimediate = MyHelper.Formatmsg("pathHOLTER=" + internalPathHOLTER.ToString());
            }
            if (e.ProgressPercentage == 29)
            {
                object Userstate = e.UserState as object;
                bool internalprintHOLTERR = (bool)e.UserState;
                this.gprintHOLTERR = internalprintHOLTERR;
                gimediate = MyHelper.Formatmsg("printHOLTERR=" + internalprintHOLTERR.ToString());
            }
            if (e.ProgressPercentage == 30)
            {
                object Userstate = e.UserState as object;
                string internalprintHOLTERR = (string)e.UserState;
                this.gHOLTERRPrinter = internalprintHOLTERR;
                gimediate = MyHelper.Formatmsg("HOLTERRPrinter=" + internalprintHOLTERR.ToString());
            }
            if (e.ProgressPercentage == 31)
            {
                object Userstate = e.UserState as object;
                bool internalprintHOLTERI = (bool)e.UserState;
                this.gprintHOLTERI = internalprintHOLTERI;
                gimediate = MyHelper.Formatmsg("printHOLTERI=" + internalprintHOLTERI.ToString());
            }
            if (e.ProgressPercentage == 32)
            {
                object Userstate = e.UserState as object;
                string internalprintHOLTERI = (string)e.UserState;
                this.gHOLTERIPrinter = internalprintHOLTERI;
                gimediate = MyHelper.Formatmsg("HOLTERIPrinter=" + internalprintHOLTERI.ToString());
            }
            if (e.ProgressPercentage == 33)
            {
                object Userstate = e.UserState as object;
                bool internalprintHOLTERC = (bool)e.UserState;
                this.gprintHOLTERC = internalprintHOLTERC;
                gimediate = MyHelper.Formatmsg("printHOLTERC=" + internalprintHOLTERC.ToString());
            }
            if (e.ProgressPercentage == 34)
            {
                object Userstate = e.UserState as object;
                string internalprintHOLTERC = (string)e.UserState;
                this.gHOLTERCPrinter = internalprintHOLTERC;
                gimediate = MyHelper.Formatmsg("HOLTERCPrinter=" + internalprintHOLTERC.ToString());
            }
            if (e.ProgressPercentage == 35)
            {
                object Userstate = e.UserState as object;
                bool internaljoinHOLTER = (bool)e.UserState;
                this.gjoinHOLTER = internaljoinHOLTER;
                gimediate = MyHelper.Formatmsg("joinHOLTER=" + internaljoinHOLTER.ToString());
            }
            if (e.ProgressPercentage == 36)
            {
                object Userstate = e.UserState as object;
                string internalprintHOLTERTbl = (string)e.UserState;
                this.gpathHOLTERTbl = internalprintHOLTERTbl;
                gimediate = MyHelper.Formatmsg("pathHOLTERTbl=" + internalprintHOLTERTbl.ToString());
            }
            if (e.ProgressPercentage == 37)
            {
                object Userstate = e.UserState as object;
                string internalHOLTERAdd = (string)e.UserState;
                this.gpathHOLTERTAdd = internalHOLTERAdd;
                gimediate = MyHelper.Formatmsg("pathHOLTERTAdd=" + internalHOLTERAdd.ToString());
            }
            if (e.ProgressPercentage == 38)
            {
                object Userstate = e.UserState as object;
                string internalHOLTERAPass = (string)e.UserState;
                this.gpathHOLTERTPass = internalHOLTERAPass;
                gimediate = MyHelper.Formatmsg("pathHOLTERTPass=" + internalHOLTERAPass.ToString());
            }
            
            
            //END HOLTER
            // START MAPA
            if (e.ProgressPercentage == 41)
            {
                object Userstate = e.UserState as object;
                string internalPathMAPA = (string)e.UserState;
                gpathMAPA = internalPathMAPA;
                gimediate = MyHelper.Formatmsg("pathMAPA=" + internalPathMAPA.ToString());
            }
            if (e.ProgressPercentage == 42)
            {
                object Userstate = e.UserState as object;
                bool internalprintMAPAR = (bool)e.UserState;
                this.gprintMAPAR = internalprintMAPAR;
                gimediate = MyHelper.Formatmsg("printMAPAR=" + internalprintMAPAR.ToString());
            }
            if (e.ProgressPercentage == 43)
            {
                object Userstate = e.UserState as object;
                string internalprintMAPAR = (string)e.UserState;
                this.gMAPARPrinter = internalprintMAPAR;
                gimediate = MyHelper.Formatmsg("MAPARPrinter=" + internalprintMAPAR.ToString());
            }
            if (e.ProgressPercentage == 44)
            {
                object Userstate = e.UserState as object;
                bool internalprintMAPAI = (bool)e.UserState;
                this.gprintMAPAI = internalprintMAPAI;
                gimediate = MyHelper.Formatmsg("printMAPAI=" + internalprintMAPAI.ToString());
            }
            if (e.ProgressPercentage == 45)
            {
                object Userstate = e.UserState as object;
                string internalprintMAPAI = (string)e.UserState;
                this.gMAPAIPrinter = internalprintMAPAI;
                gimediate = MyHelper.Formatmsg("MAPAIPrinter=" + internalprintMAPAI.ToString());
            }
            if (e.ProgressPercentage == 46)
            {
                object Userstate = e.UserState as object;
                bool internalprintMAPAC = (bool)e.UserState;
                this.gprintMAPAC = internalprintMAPAC;
                gimediate = MyHelper.Formatmsg("printMAPAC=" + internalprintMAPAC.ToString());
            }
            if (e.ProgressPercentage == 47)
            {
                object Userstate = e.UserState as object;
                string internalprintMAPAC = (string)e.UserState;
                this.gMAPACPrinter = internalprintMAPAC;
                gimediate = MyHelper.Formatmsg("MAPACPrinter=" + internalprintMAPAC.ToString());
            }
            if (e.ProgressPercentage == 48)
            {
                object Userstate = e.UserState as object;
                bool internaljoinMAPA = (bool)e.UserState;
                this.gjoinMAPA = internaljoinMAPA;
                gimediate = MyHelper.Formatmsg("joinMAPA=" + internaljoinMAPA.ToString());
            }
            if (e.ProgressPercentage == 49)
            {
                object Userstate = e.UserState as object;
                string internalprintMAPATbl = (string)e.UserState;
                this.gpathMAPATbl = internalprintMAPATbl;
                gimediate = MyHelper.Formatmsg("pathMAPATbl=" + internalprintMAPATbl.ToString());
            }
            //END MAPA
            //START PE
            if (e.ProgressPercentage == 50)
            {
                object Userstate = e.UserState as object;
                string internalPathPE = (string)e.UserState;
                gpathPE = internalPathPE;
                gimediate = MyHelper.Formatmsg("pathPE=" + internalPathPE.ToString());
            }
            if (e.ProgressPercentage == 51)
            {
                object Userstate = e.UserState as object;
                bool internalprintPER = (bool)e.UserState;
                this.gprintPER = internalprintPER;
                gimediate = MyHelper.Formatmsg("printPER=" + internalprintPER.ToString());
            }
            if (e.ProgressPercentage == 52)
            {
                object Userstate = e.UserState as object;
                string internalprintPER = (string)e.UserState;
                this.gPERPrinter = internalprintPER;
                gimediate = MyHelper.Formatmsg("PERPrinter=" + internalprintPER.ToString());
            }
            if (e.ProgressPercentage == 53)
            {
                object Userstate = e.UserState as object;
                bool internalprintPEI = (bool)e.UserState;
                this.gprintPEI = internalprintPEI;
                gimediate = MyHelper.Formatmsg("printPEI=" + internalprintPEI.ToString());
            }
            if (e.ProgressPercentage == 54)
            {
                object Userstate = e.UserState as object;
                string internalprintPEI = (string)e.UserState;
                this.gPEIPrinter = internalprintPEI;
                gimediate = MyHelper.Formatmsg("PEIPrinter=" + internalprintPEI.ToString());
            }
            if (e.ProgressPercentage == 55)
            {
                object Userstate = e.UserState as object;
                bool internalprintPEC = (bool)e.UserState;
                this.gprintPEC = internalprintPEC;
                gimediate = MyHelper.Formatmsg("printPEC=" + internalprintPEC.ToString());
            }
            if (e.ProgressPercentage == 56)
            {
                object Userstate = e.UserState as object;
                string internalprintPEC = (string)e.UserState;
                this.gPECPrinter = internalprintPEC;
                gimediate = MyHelper.Formatmsg("PECPrinter=" + internalprintPEC.ToString());
            }
            if (e.ProgressPercentage == 57)
            {
                object Userstate = e.UserState as object;
                bool internaljoinPE = (bool)e.UserState;
                this.gjoinPE = internaljoinPE;
                gimediate = MyHelper.Formatmsg("joinPE=" + internaljoinPE.ToString());
            }
            if (e.ProgressPercentage == 58)
            {
                object Userstate = e.UserState as object;
                string internalprintPETbl = (string)e.UserState;
                this.gpathPETbl = internalprintPETbl;
                gimediate = MyHelper.Formatmsg("pathPETbl=" + internalprintPETbl.ToString());
            }
            if (e.ProgressPercentage == 59)
            {
                object Userstate = e.UserState as object;
                string internalpathUnconfirmed = (string)e.UserState;
                this.gpathUnconfirmed = internalpathUnconfirmed;
                gimediate = MyHelper.Formatmsg("pathUnconfirmed=" + internalpathUnconfirmed.ToString());
            }
            if (e.ProgressPercentage == 60)
            {
                object Userstate = e.UserState as object;
                string internalpathstorereports = (string)e.UserState;
                this.gStoreReports  = internalpathstorereports;
                gimediate = MyHelper.Formatmsg("StoreReports=" + internalpathstorereports.ToString());
            }
            if (e.ProgressPercentage == 61)
            {
                object Userstate = e.UserState as object;
                string internalpathstoreimages = (string)e.UserState;
                this.gStoreImages = internalpathstoreimages;
                gimediate = MyHelper.Formatmsg("StoreImages=" + internalpathstoreimages.ToString());
            }
            if (e.ProgressPercentage == 62)
            {
                object Userstate = e.UserState as object;
                string internalpathstorecomplete = (string)e.UserState;
                this.gStoreComplete = internalpathstorecomplete;
                gimediate = MyHelper.Formatmsg("StoreComplete=" + internalpathstorecomplete.ToString());
            }
            if (e.ProgressPercentage == 63)
            {
                object Userstate = e.UserState as object;
                string internalEmailInc = (string)e.UserState;
                this.gEmailInc = internalEmailInc;
                gimediate = MyHelper.Formatmsg("HolterEmailPath=" + internalEmailInc.ToString());
            }
            if (e.ProgressPercentage == 64)
            {
                object Userstate = e.UserState as object;
                string internalPdfPrinter = (string)e.UserState;
                this.gPdfPrinter = internalPdfPrinter;
                gimediate = MyHelper.Formatmsg("PdfPrinter=" + internalPdfPrinter.ToString());
            } 
            #endregion

        }
        private void BCKINIReader_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            object[] resultados = e.Result as object[];
            try
            {
                glastinicheck = (DateTime)resultados[0];
                gimediate = MyHelper.Formatmsg("Ficheiro INI analisado.");
                timerUpdateLst.Enabled = true;
                if (gstate == 0)
                {
                    this.timerFlush.Enabled = true;
                    this.timerFolders.Enabled = true;
                    this.timerEmail.Enabled = true;
                    this.frmdebug.starttimer();
                }
            }
            catch
            {

                // definir erro
            }
            
            
        }
        private void BCKMasterWriter_DoWork(object sender, DoWorkEventArgs e)
        {
            List<string> erros = new List<string>();
            int error = 0;
            object[] parameters = e.Argument as object[];
            string InternalLogPath = @"C:\PriJobAide\Logs";
            Helper internalhelp = new Helper();
            ListLogger MWtowrite = (ListLogger)parameters[0];
            try
            {
                InternalLogPath = (string)parameters[1];
            }
            catch
            {
                error++;
                erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir parametros em BCKMasterWriter.")); 
            }
            if (Directory.Exists(InternalLogPath) == false)
            {
                try
                {
                    Directory.CreateDirectory(InternalLogPath);
                }
                catch
                {
                    error++;
                    erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao criar directoria " + InternalLogPath));
                }
            }
            string file = InternalLogPath + @"\" + DateTime.Now.ToShortDateString() + ".txt";
            if (File.Exists(file))
            {
                bool fileisinuse = internalhelp.FileInUse(file);

                if (fileisinuse == true)
                {

                    DateTime quit = DateTime.Now;
                    while (fileisinuse == true)
                    {
                        Thread.Sleep(500);
                        fileisinuse = internalhelp.FileInUse(file);
                        if (quit.AddSeconds(5) < DateTime.Now)
                        {
                            error++;
                            erros.Add(internalhelp.Formatmsg("Não foi possível aceder a " + file));
                            break;
                        }
                    }
                }
                
                if (error == 0)
                {
                    MWtowrite.AddEx(erros);
                    using (StreamWriter sw = File.AppendText(file))
                    {
                        MWtowrite.GetList.ForEach(sw.WriteLine);
                        sw.Close();
                    }
                    MWtowrite.RemoveEx(MWtowrite.GetList);
                }
            }
            else
            {
                try
                {
                    using (StreamWriter sw = File.CreateText(file))
                    {
                        MWtowrite.GetList.ForEach(sw.WriteLine);
                        sw.Close();
                    }
                    MWtowrite.RemoveEx(MWtowrite.GetList);
                }
                catch {  }
            }
        }
        private void BCKFlush_DoWork(object sender, DoWorkEventArgs e)
        {
            List<string> erros = new List<string>();
            int error = 0;
            
            Helper internalhelp = new Helper();
            object[] parameters = e.Argument as object[];
            ListLogger Mwtowrite = (ListLogger)parameters[0];
            string internalPathLogs = "";
            string internalPathECG = "";
            int internalgDaystoFlush = 10;
            string internalPathECO = "";
            string internalPathMAPA = "";
            string internalPathHOLTER = "";
            string internalPathPE = "";
            string internalPathPDF = "";
            string internalStoreReports = "";
            string internalStoreImages = "";
            string internalStoreComplete = "";
            DateTime internalLastFlush = (DateTime)parameters[9];
            string internalpathUnconfirmed = "";
            int result = 0;
            Thread.Sleep(1000);
            try
            {
                 internalStoreComplete = (string)parameters[13];
                 
            }
            catch { error++;  erros.Add (internalhelp.Formatmsg ("Ocorreu um erro ao atribuir valores a internalStoreComplete em BCKFlush."));}
            try
            {
                 internalStoreImages = (string)parameters[12];
                 
            }
            catch { error++;  erros.Add (internalhelp.Formatmsg ("Ocorreu um erro ao atribuir valores a internalStoreImages em BCKFlush."));}
            try
            {
                 internalStoreReports = (string)parameters[11];
                 
            }
            catch { error++;  erros.Add (internalhelp.Formatmsg ("Ocorreu um erro ao atribuir valores a internalStoreReports em BCKFlush."));}
            try
            {
                 internalpathUnconfirmed = (string)parameters[10];
                 
            }
            catch { error++;  erros.Add (internalhelp.Formatmsg ("Ocorreu um erro ao atribuir valores a internalpathUnconfirmed em BCKFlush."));}
            try
            {
                 internalPathLogs = (string)parameters[1];
                 
            }
            catch { error++;  erros.Add (internalhelp.Formatmsg ("Ocorreu um erro ao atribuir valores a internalPathLogs em BCKFlush."));}
            try
            {
                internalPathECG = (string)parameters[2];

            }
            catch { error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathECG em BCKFlush.")); }
            try
            {
                internalgDaystoFlush = (int)parameters[8];

            }
            catch { error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalgDaystoFlush em BCKFlush.")); }
            try
            {
                internalPathECO = (string)parameters[3];

            }
            catch { error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalPathECO em BCKFlush.")); }
            try
            {
                internalPathHOLTER = (string)parameters[4];

            }
            catch { error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalPathHOLTER em BCKFlush.")); }
            try
            {
                internalPathMAPA = (string)parameters[5];

            }
            catch { error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalPathMAPA em BCKFlush.")); }
            try
            {
                internalPathPE = (string)parameters[6];

            }
            catch { error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalPathPE em BCKFlush.")); }
            try
            {
                internalPathPDF = (string)parameters[7];

            }
            catch { error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalPathPDF em BCKFlush.")); }
            try
            {
                internalLastFlush = (DateTime)parameters[9];

            }
            catch { error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalLastFlush em BCKFlush.")); }
            if (error == 0)
            {
                if (internalLastFlush.AddHours(1) < DateTime.Now)
                {
                    List<string> foldersToerase = new List<string>();
                    List<string> filesToerase = new List<string>();
                    filesToerase.Add(internalPathLogs);
                    filesToerase.Add(internalPathECG);
                    filesToerase.Add(internalPathECO);
                    filesToerase.Add(internalPathHOLTER);
                    filesToerase.Add(internalPathMAPA);
                    filesToerase.Add(internalPathPE);
                    filesToerase.Add(internalpathUnconfirmed);
                    foldersToerase.Add(internalStoreReports);
                    foldersToerase.Add(internalStoreImages);
                    foldersToerase.Add(internalStoreComplete);

                   
                    if (foldersToerase.Count > 0)
                    {
                        foreach (string sf in foldersToerase)
                        {
                            List<string> subDirectories = new List<string>();
                            if (Directory.Exists(sf))
                            { 
                            subDirectories.AddRange(Directory.GetDirectories(sf));
                            }
                            if (subDirectories.Count > 0)
                            {

                                foreach (string dirs in subDirectories)
                                    {


                                        if (Directory.Exists(dirs))
                                        {

                                            DirectoryInfo a = new DirectoryInfo(dirs);
                                            if (a.CreationTime.AddDays(internalgDaystoFlush) < DateTime.Now)
                                            {
                                                try
                                                {
                                                    a.Delete (true); erros.Add(internalhelp.Formatmsg("A directoria " + a.Name + " foi apagada com sucesso.")); BCKFlush.ReportProgress(0, a.Name);
                                                }
                                                catch { erros.Add(internalhelp.Formatmsg("Não foi possível apagar a directoria " + a.Name + ".")); }
                                            }
                                        }
                                    }
                                
                            }
                        }
                    
                    }
                    if (filesToerase.Count > 0)
                    {

                        foreach (string s in filesToerase)
                        {

                            if (Directory.Exists(s))
                            {

                                


                                
                                List<string> FileInDir = new List<string>();
                                FileInDir = Directory.GetFiles(s).ToList();
                                if (FileInDir.Count > 0)
                                {

                                    foreach (string files in FileInDir)
                                    {
                                        if (File.Exists(files))
                                        {
                                            FileInfo b = new FileInfo(files);
                                            if (b.CreationTime.AddDays(internalgDaystoFlush) < DateTime.Now)
                                            {
                                                try { b.Delete(); erros.Add(internalhelp.Formatmsg("O ficheiro " + b.Name + " foi apagado com sucesso.")); BCKFlush.ReportProgress(1, b.Name); }
                                                catch { erros.Add(internalhelp.Formatmsg("Não foi possível apagar o ficheiro " + b.Name + ".")); }
                                            }
                                        }
                                    }
                                }
                                

                            }
                            else { erros.Add(internalhelp.Formatmsg("A directoria " + s + " não existe.")); }

                        }
                    }
                    result = 0;
                }
                else
                {
                    result = 1;
                }
            }
            object[] resultados = new object[] { result };
            e.Result = resultados;
            if (erros.Count > 0)
            {
                Mwtowrite.AddEx(erros);
            }
            
        }
        private void BCKFlush_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            object[] parameters = e.Result as object[];
            int result = (int)parameters[0];
            if (result == 0)
            { gLastFlush = DateTime.Now; gimediate = MyHelper.Formatmsg("Ultimo Flush a: " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")); }
            else { }
            
            
        }
        private void BCKFlush_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 0)
            {
                gimediate = MyHelper.Formatmsg("A directoria " + (string)e.UserState + " foi apagada com sucesso.");
            }
            if (e.ProgressPercentage == 1)
            { gimediate = MyHelper.Formatmsg("O ficheiro " + (string)e.UserState + " foi apagado com sucesso."); }
        }       
        private void BCKFolders_DoWork(object sender, DoWorkEventArgs e)
        {
            Helper Myhelp = new Helper();

            #region Variaveis Globais
            List<string> erros = new List<string>();
            int error = 0;
            object[] parameters = e.Argument as object[];
            ListLogger internallogger = (ListLogger)parameters[0];
            string internalpathECG = "";
            string internalpathECO = "";
            string internalpathHOLTER = "";
            string internalpathMAPA = "";
            string internalpathPE = "";
            string internalpathUC = "";
            bool internaldebug = false;

            string internalBookletAudit = "";
            string internalBookletPrinter = "";
            string internalBookletProcess = "";
            int internalBookletSecsToWait = 30;
            DateTime internalLastBookletPrint = DateTime.Now;
            string internalSumatraPdfpath = "";

            string internalECGRPrinter = "";
            string internalECGIPrinter = "";
            string internalECGCPrinter = "";

            bool internalECGRPrint = false;
            bool internalECGIPrint = false;
            bool internalECGCPrint = false;
            bool internaljoinECG = false;

            string internalECORPrinter = "";
            string internalECOIPrinter = "";
            string internalECOCPrinter = "";

            bool internalECORPrint = false;
            bool internalECOIPrint = false;
            bool internalECOCPrint = false;
            bool internaljoinECO = false;

            string internalHOLTERRPrinter = "";
            string internalHOLTERIPrinter = "";
            string internalHOLTERCPrinter = "";

            bool internalHOLTERRPrint = false;
            bool internalHOLTERIPrint = false;
            bool internalHOLTERCPrint = false;
            bool internaljoinHOLTER = false;

            string internalMAPARPrinter = "";
            string internalMAPAIPrinter = "";
            string internalMAPACPrinter = "";

            bool internalMAPARPrint = false;
            bool internalMAPAIPrint = false;
            bool internalMAPACPrint = false;
            bool internaljoinMAPA = false;

            string internalPERPrinter = "";
            string internalPEIPrinter = "";
            string internalPECPrinter = "";

            bool internalPERPrint = false;
            bool internalPEIPrint = false;
            bool internalPECPrint = false;
            bool internaljoinPE = false;

            string internalstorereports = "";
            string internalstoreimages = "";
            string internalstorecomplete = "";
            try
            {
                internalSumatraPdfpath = (string)parameters[46];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalSumatraPdfpath em BCKFolders.")); }
            #region ECG
            try
            {
                internaljoinECG = (bool)parameters[33];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaljoinECG em BCKFolders.")); }
            try
            {
                internalECGRPrint = (bool)parameters[18];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalECGRPrint em BCKFolders.")); }
            try
            {
                internalECGIPrint = (bool)parameters[19];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalECGIPrint em BCKFolders.")); }
            try
            {
                internalECGCPrint = (bool)parameters[20];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalECGCPrint em BCKFolders.")); }
            try
            {
                internalECGRPrinter = (string)parameters[6];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalECGRPrinter em BCKFolders.")); }
            try
            {
                internalECGIPrinter = (string)parameters[7];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalECGIPrinter em BCKFolders.")); }
            try
            {
                internalECGCPrinter = (string)parameters[8];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalECGCPrinter em BCKFolders.")); }
            #endregion
            #region ECO
            try
            {
                internaljoinECO = (bool)parameters[34];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaljoinECO em BCKFolders.")); }
            try
            {
                internalECORPrint = (bool)parameters[21];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalECORPrint em BCKFolders.")); }
            try
            {
                internalECOIPrint = (bool)parameters[22];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalECOIPrint em BCKFolders.")); }
            try
            {
                internalECOCPrint = (bool)parameters[23];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalECOCPrint em BCKFolders.")); }
            try
            {
                internalECORPrinter = (string)parameters[9];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalECORPrinter em BCKFolders.")); }
            try
            {
                internalECOIPrinter = (string)parameters[10];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalECOIPrinter em BCKFolders.")); }
            try
            {
                internalECOCPrinter = (string)parameters[11];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalECOCPrinter em BCKFolders.")); }
            #endregion
            #region HOLTER
            try
            {
                internaljoinHOLTER = (bool)parameters[35];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaljoinHOLTER em BCKFolders.")); }
            try
            {
                internalHOLTERRPrint = (bool)parameters[24];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalHOLTERRPrint em BCKFolders.")); }
            try
            {
                internalHOLTERIPrint = (bool)parameters[25];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalHOLTERIPrint em BCKFolders.")); }
            try
            {
                internalHOLTERCPrint = (bool)parameters[26];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalHOLTERCPrint em BCKFolders.")); }
            try
            {
                internalHOLTERRPrinter = (string)parameters[12];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalHOLTERRPrinter em BCKFolders.")); }
            try
            {
                internalHOLTERIPrinter = (string)parameters[13];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalHOLTERIPrinter em BCKFolders.")); }
            try
            {
                internalHOLTERCPrinter = (string)parameters[14];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalHOLTERCPrinter em BCKFolders.")); }
            #endregion
            #region MAPA
            try
            {
                internaljoinMAPA = (bool)parameters[36];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaljoinMAPA em BCKFolders.")); }
            try
            {
                internalMAPARPrint = (bool)parameters[27];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalMAPARPrint em BCKFolders.")); }
            try
            {
                internalMAPAIPrint = (bool)parameters[28];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalMAPAIPrint em BCKFolders.")); }
            try
            {
                internalMAPACPrint = (bool)parameters[29];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalMAPACPrint em BCKFolders.")); }
            try
            {
                internalMAPARPrinter = (string)parameters[15];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalMAPARPrinter em BCKFolders.")); }
            try
            {
                internalMAPAIPrinter = (string)parameters[16];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalMAPAIPrinter em BCKFolders.")); }
            try
            {
                internalMAPACPrinter = (string)parameters[17];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalMAPACPrinter em BCKFolders.")); }
            #endregion
            #region PE
            try
            {
                internaljoinPE = (bool)parameters[37];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaljoinPE em BCKFolders.")); }
            try
            {
                internalPERPrint = (bool)parameters[30];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalPERPrint em BCKFolders.")); }
            try
            {
                internalPEIPrint = (bool)parameters[31];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalPEIPrint em BCKFolders.")); }
            try
            {
                internalPECPrint = (bool)parameters[32];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalPECPrint em BCKFolders.")); }
            try
            {
                internalPERPrinter = (string)parameters[43];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalPERPrinter em BCKFolders.")); }
            try
            {
                internalPEIPrinter = (string)parameters[44];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalPEIPrinter em BCKFolders.")); }
            try
            {
                internalPECPrinter = (string)parameters[45];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalPECPrinter em BCKFolders.")); }
            #endregion
            #region Geral
            try
            {
                internalpathECG = (string)parameters[1];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathECG em BCKFolders.")); }
            try
            {
                internalpathECO = (string)parameters[2];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathECO em BCKFolders.")); }
            try
            {
                internalpathHOLTER = (string)parameters[3];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathHOLTER em BCKFolders.")); }
            try
            {
                internalpathMAPA = (string)parameters[4];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathMAPA em BCKFolders.")); }
            try
            {
                internalpathPE = (string)parameters[5];

            }

            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathPE em BCKFolders.")); }
            try
            {
                internalBookletPrinter = (string)parameters[38];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalBookletPrinter em BCKFolders.")); }
            try
            {
                internalBookletAudit = (string)parameters[39];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalBookletAudit em BCKFolders.")); }
            try
            {
                internalBookletProcess = (string)parameters[40];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalBookletProcess em BCKFolders.")); }
            try
            {
                internalBookletSecsToWait = (int)parameters[41];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalBookletSecsToWait em BCKFolders.")); }
            try
            {
                internalLastBookletPrint = (DateTime)parameters[42];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalLastBookletPrint em BCKFolders.")); }
            try
            {
                internalstorereports = (string)parameters[47];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalstorereports em BCKFolders.")); }
            try
            {
                internalstoreimages = (string)parameters[48];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalstoreimages em BCKFolders.")); }
            try
            {
                internalstorecomplete = (string)parameters[49];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalstorecomplete em BCKFolders.")); }
            try
            {
                internalpathUC = (string)parameters[50];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathUC em BCKFolders.")); }
            try
            {
                internaldebug = (bool)parameters[51];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaldebug em BCKFolders.")); }
            #endregion 
            #endregion
            if (error == 0)
            {
                List<string> analyzingpatf = new List<string>();
                
                analyzingpatf.Add(internalpathECG);
                analyzingpatf.Add(internalpathECO);
                analyzingpatf.Add(internalpathHOLTER);
                analyzingpatf.Add(internalpathMAPA);
                analyzingpatf.Add(internalpathPE);
                analyzingpatf.Add(internalpathUC);
                #region Associações
                for (int i = 0; i <= analyzingpatf.Count - 1; i++)
                {
                    Thread.Sleep(500);
                    string pathtoanalyze = analyzingpatf[i];
                    string internalTempRPrinter = "";
                    string internalTempIPrinter = "";
                    string internalTempCPrinter = "";

                    bool internalTempRPrint = false;
                    bool internalTempIPrint = false;
                    bool internalTempCPrint = false;
                    bool internalTempjoin = false;
                    if (i == 0)
                    {
                        internalTempRPrinter = internalECGRPrinter;
                        internalTempIPrinter = internalECGIPrinter;
                        internalTempCPrinter = internalECGCPrinter;
                        internalTempRPrint = internalECGRPrint;
                        internalTempIPrint = internalECGIPrint;
                        internalTempCPrint = internalECGCPrint;
                        internalTempjoin = internaljoinECG;
                    }
                    if (i == 1)
                    {
                        internalTempRPrinter = internalECORPrinter;
                        internalTempIPrinter = internalECOIPrinter;
                        internalTempCPrinter = internalECOCPrinter;
                        internalTempRPrint = internalECORPrint;
                        internalTempIPrint = internalECOIPrint;
                        internalTempCPrint = internalECOCPrint;
                        internalTempjoin = internaljoinECO;
                    }
                    if (i == 2)
                    {
                        internalTempRPrinter = internalHOLTERRPrinter;
                        internalTempIPrinter = internalHOLTERIPrinter;
                        internalTempCPrinter = internalHOLTERCPrinter;
                        internalTempRPrint = internalHOLTERRPrint;
                        internalTempIPrint = internalHOLTERIPrint;
                        internalTempCPrint = internalHOLTERCPrint;
                        internalTempjoin = internaljoinHOLTER;
                    }
                    if (i == 3)
                    {
                        internalTempRPrinter = internalMAPARPrinter;
                        internalTempIPrinter = internalMAPAIPrinter;
                        internalTempCPrinter = internalMAPACPrinter;
                        internalTempRPrint = internalMAPARPrint;
                        internalTempIPrint = internalMAPAIPrint;
                        internalTempCPrint = internalMAPACPrint;
                        internalTempjoin = internaljoinMAPA;
                    }
                    if (i == 4)
                    {
                        internalTempRPrinter = internalPERPrinter;
                        internalTempIPrinter = internalPEIPrinter;
                        internalTempCPrinter = internalPECPrinter;
                        internalTempRPrint = internalPERPrint;
                        internalTempIPrint = internalPEIPrint;
                        internalTempCPrint = internalPECPrint;
                        internalTempjoin = internaljoinPE;
                    }
                #endregion
                    if (Directory.Exists(pathtoanalyze))
                    {
                        List<Prijob> Jobs = new List<Prijob>();
                        List<string> PathFiles = Directory.GetFiles(pathtoanalyze,"*.txt", SearchOption.TopDirectoryOnly).ToList();
                        if (PathFiles.Count > 0)
                        {
                            foreach (string txtfile in PathFiles)
                            #region Criar Jobs
                            {

                                string associatePdf = txtfile.Replace(".txt", ".pdf");
                                if (File.Exists(associatePdf))
                                {
                                    FileInfo Pdf = new FileInfo(associatePdf);
                                    if (Pdf.Length > 0)
                                        {
                                            
                                                List<string> content = new List<string>();
                                                try
                                                {
                                                    content = File.ReadAllLines(txtfile).ToList();
                                                }
                                                catch { erros.Add(Myhelp.Formatmsg("Não foi possível ler o ficheiro " + txtfile)); }//mensagem de erro na leitura do ficheiro contente
                                                if (content.Count >= 8)
                                                {
                                                    if (Jobs.Count == 0)
                                                    {
                                                        Prijob newjob = new Prijob(content[0], content[1]);
                                                        if (content[2] == "R")
                                                        { newjob.AddReport(content); }
                                                        if (content[2] == "I")
                                                        { newjob.AddImage(content); }
                                                        if (content[2] == "C")
                                                        { newjob.AddComplete(content); }
                                                        Jobs.Add(newjob);
                                                    }
                                                    else
                                                    {
                                                        if (Jobs.Any(job => job.GId == content[1]))
                                                        {
                                                            foreach (Prijob Job in Jobs)
                                                            {
                                                                if (Job.GId == content[1])
                                                                {
                                                                    if (content[2] == "R")
                                                                    { Job.AddReport(content); }
                                                                    if (content[2] == "I")
                                                                    { Job.AddImage(content); }
                                                                    if (content[2] == "C")
                                                                    { Job.AddComplete(content); }
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Prijob newjob = new Prijob(content[0], content[1]);
                                                            if (content[2] == "R")
                                                            { newjob.AddReport(content); }
                                                            if (content[2] == "I")
                                                            { newjob.AddImage(content); }
                                                            if (content[2] == "C")
                                                            { newjob.AddComplete(content); }
                                                            Jobs.Add(newjob);
                                                        }
                                                    }
                                                }
                                                else { erros.Add(Myhelp.Formatmsg(txtfile + " é um ficheiro inválido.")); }
                                            
                                        }
                                        else { erros.Add(Myhelp.Formatmsg(associatePdf + " é invalido ou está corrompido.")); }
                                    
                                    
                                }
                                else { try { File.Delete(txtfile); } catch { erros.Add(Myhelp.Formatmsg("Não foi possível apagar " + txtfile)); } erros.Add(Myhelp.Formatmsg(associatePdf + " não existe.")); }
                            } 
                            #endregion
                        }
                        
                        //primeiro encontrar associações
                        foreach (Prijob Job in Jobs)
                        {
                            List<List<string>> rep = Job.GReports;
                            List<List<string>> img = Job.GImages;
                            List<List<string>> cpt = Job.GComplete;
                            foreach (List<string> rp in rep)
                            {
                                #region Caso Armazenado
                                if (rp[4].Contains("Armazenado."))
                                {
                                    if (rp[8] != "default")
                                    {
                                        FileInfo a = new FileInfo(pathtoanalyze + @"\" + rp[6] + ".pdf");
                                        FileInfo b = new FileInfo(pathtoanalyze + @"\" + rp[6] + ".txt");
                                        if (a.CreationTime.AddHours(4) < DateTime.Now)
                                        {
                                            if (Myhelp.FileInUse(a.FullName) == false && Myhelp.FileInUse(b.FullName) == false)
                                            {
                                                try { File.Delete(a.FullName); File.Delete(b.FullName); }
                                                catch { erros.Add(Myhelp.Formatmsg("Não foi possível apagar " + a.FullName)); }//erro ao apagar ficheiros
                                            }
                                        }
                                    }
                                } 
                                #endregion
                                #region Caso Aguarda armazenamento
                                if (rp[4].Contains("Aguarda armazenamento."))
                                {
                                    if (rp[3] != "default")
                                    {
                                        bool inerror = false;
                                        if (Directory.Exists(internalstorereports + @"\" + rp[3]) == false)
                                        { Directory.CreateDirectory(internalstorereports + @"\" + rp[3]); Thread.Sleep(500); }
                                        try { File.Copy(pathtoanalyze + @"\" + rp[6] + ".pdf", internalstorereports + @"\" + rp[3] + @"\" + rp[5] + "." + rp[1] + ".pdf", true); File.Copy(pathtoanalyze + @"\" + rp[6] + ".txt", internalstorereports + @"\" + rp[3] + @"\" + rp[5] + "." + rp[1] + ".txt", true); }
                                        catch { inerror = true; erros.Add(Myhelp.Formatmsg("Não foi possível copiar " + pathtoanalyze + @"\" + rp[6] + ".pdf")); }//mensagem de erro ao copiar
                                        if (inerror == false)
                                        { rp[4] = "Armazenado."; }
                                    }
                                } 
                                #endregion
                                #region Aguarda Imagem
                                if (rp[4].Contains("Aguarda Imagem."))
                                {
                                    if (internalTempjoin == true)
                                    {
                                        if (rp[7] == "default" && rp[8] == "default")
                                        {
                                            foreach (List<string> im in img)
                                            {
                                                if (im[4].Contains("Aguarda Relatorio."))
                                                {
                                                    if (im[6] == "default" && im[8] == "default")
                                                    {
                                                        if (im[3] == "default" && rp[3] != "default")
                                                        { im[3] = rp[3]; }
                                                        List<string> newcomplete = new List<string>();
                                                        newcomplete.Add(rp[0]);
                                                        newcomplete.Add(rp[1]);
                                                        newcomplete.Add("C");
                                                        newcomplete.Add(rp[3]);
                                                        newcomplete.Add("Aguarda impressao.");
                                                        if (pathtoanalyze == internalpathECG)
                                                        { newcomplete.Add("ECG"); }
                                                        if (pathtoanalyze == internalpathECO)
                                                        { newcomplete.Add("ECO"); }
                                                        if (pathtoanalyze == internalpathMAPA)
                                                        { newcomplete.Add("MAPA"); }
                                                        if (pathtoanalyze == internalpathHOLTER)
                                                        { newcomplete.Add("HOLTER"); }
                                                        if (pathtoanalyze == internalpathPE)
                                                        { newcomplete.Add("PE"); }
                                                        newcomplete.Add(rp[6]);
                                                        newcomplete.Add(im[7]);
                                                        string targetfile = pathtoanalyze + @"\" + rp[5] + "C";
                                                        int x = 0;
                                                        if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                        {
                                                            while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                            { x++; }
                                                        }
                                                        string completefilename = targetfile + x.ToString() + ".pdf";
                                                        newcomplete.Add(rp[5] + "C" + x.ToString());
                                                        if (Myhelp.FileInUse(pathtoanalyze + @"\" + rp[6] + ".pdf") == false && Myhelp.FileInUse(pathtoanalyze + @"\" + im[7] + ".pdf") == false)
                                                        {
                                                            string[] files = new string[] { pathtoanalyze + @"\" + rp[6] + ".pdf", pathtoanalyze + @"\" + im[7] + ".pdf" };
                                                            bool sucess = Myhelp.MergePDF(completefilename, files);
                                                            if (sucess == true)
                                                            {
                                                                rp[8] = rp[5] + "C" + x.ToString();
                                                                rp[7] = im[7];
                                                                rp[4] = "Aguarda armazenamento.";
                                                                im[8] = rp[5] + "C" + x.ToString();
                                                                im[6] = rp[6];
                                                                im[4] = "Aguarda armazenamento.";
                                                                cpt.Add(newcomplete);
                                                            }
                                                        }
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (rp[7] == "default")
                                        {
                                            rp[7] = "false";
                                        }
                                        //criar complete com cp[7] = rp[7}
                                        List<string> newcomplete = new List<string>();
                                        newcomplete.Add(rp[0]);
                                        newcomplete.Add(rp[1]);
                                        newcomplete.Add("C");
                                        newcomplete.Add(rp[3]);
                                        newcomplete.Add("Aguarda impressao.");
                                        if (pathtoanalyze == internalpathECG)
                                        { newcomplete.Add("ECG"); }
                                        if (pathtoanalyze == internalpathECO)
                                        { newcomplete.Add("ECO"); }
                                        if (pathtoanalyze == internalpathMAPA)
                                        { newcomplete.Add("MAPA"); }
                                        if (pathtoanalyze == internalpathHOLTER)
                                        { newcomplete.Add("HOLTER"); }
                                        if (pathtoanalyze == internalpathPE)
                                        { newcomplete.Add("PE"); }
                                        newcomplete.Add(rp[6]);
                                        newcomplete.Add(rp[7]);
                                        string targetfile = pathtoanalyze + @"\" + rp[5] + "C";
                                        int x = 0;
                                        if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                        {
                                            while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                            { x++; }
                                        }
                                        string completefilename = targetfile + x.ToString() + ".pdf";
                                        newcomplete.Add(rp[5] + "C" + x.ToString());
                                        rp[8] = rp[5] + "C" + x.ToString();
                                        if (Myhelp.FileInUse(pathtoanalyze + @"\" + rp[6] + ".pdf") == false)
                                        {

                                            bool sucess = false;
                                            try { File.Copy(pathtoanalyze + @"\" + rp[6] + ".pdf", completefilename); sucess = true; }
                                            catch { sucess = false; }
                                            if (sucess == true)
                                            {
                                                rp[8] = rp[5] + "C" + x.ToString();
                                                rp[4] = "Aguarda armazenamento.";
                                                cpt.Add(newcomplete);
                                            }
                                        }
                                    }
                                } 
                                #endregion
                                #region Caso Impresso
                                if (rp[4].Contains("Impresso."))
                                {
                                    rp[4] = "Aguarda Imagem.";
                                } 
                                #endregion
                                #region Caso Aguarda impressao
                                if (rp[4].Contains("Aguarda impressao."))
                                {
                                    bool toprint = internalTempRPrint;
                                    string printer = internalTempRPrinter;
                                    FileInfo a = new FileInfo(pathtoanalyze + @"\" + rp[6] + ".pdf");

                                    if (File.Exists(internalSumatraPdfpath))
                                    {
                                        if (toprint == true)
                                        {
                                            if (Myhelp.printerisvalid(printer))
                                            {
                                                if (printer == internalBookletPrinter)//Bookletprint
                                                {
                                                    bool procrunning = Myhelp.isProcessRunning(internalBookletProcess);
                                                    if (procrunning == false)
                                                    {
                                                        if (Directory.Exists(internalBookletAudit))
                                                        {
                                                            string AuditIdentifier = "Print-in." + DateTime.Now.ToString("yyyyMMdd") + ".xml";
                                                            if (File.Exists(internalBookletAudit + @"\" + AuditIdentifier))
                                                            {
                                                                if (internaldebug == true)
                                                                {
                                                                    erros.Add(Myhelp.Formatmsg("Última impressão a :" + Myhelp.LastBookletPrintTime(internalBookletAudit).ToString() + " Com SecsTowait - " + Myhelp.LastBookletPrintTime(internalBookletAudit).AddSeconds(internalBookletSecsToWait).ToString() + " Now: " + DateTime.Now.ToString()));
                                                                }
                                                                if (Myhelp.LastBookletPrintTime(internalBookletAudit).AddSeconds(internalBookletSecsToWait) < DateTime.Now)
                                                                {
                                                                    Process p = new System.Diagnostics.Process();
                                                                    ProcessStartInfo startInfo = new ProcessStartInfo();
                                                                    p.StartInfo.FileName = internalSumatraPdfpath;
                                                                    if (internaldebug == true)
                                                                    { p.StartInfo.Arguments = "-exit-when-done -print-to \"" + printer + "\" " + a.FullName; }
                                                                    else
                                                                    {
                                                                        p.StartInfo.Arguments = "-silent -exit-when-done -print-to \"" + printer + "\" " + a.FullName;
                                                                    }
                                                                    p.Start();
                                                                    p.WaitForExit();
                                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " foi enviado para Booklet."));
                                                                    Thread.Sleep(500);
                                                                    //MessageBox.Show(BookletAudit + " " + a.Name + DateTime.Now.ToString() + " " + internalBookletSecsToWait.ToString());
                                                                    bool FoundPrint = Myhelp.FoundPrintJob(BookletAudit, a.Name, DateTime.Now, internalBookletSecsToWait);
                                                                    if (FoundPrint == true)
                                                                    {
                                                                        erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " foi encontrado em " + AuditIdentifier));
                                                                        rp[4] = "Impresso.";
                                                                    }
                                                                    else
                                                                    {
                                                                        erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " não foi encontrado em " + AuditIdentifier));
                                                                        rp[4] = "Impresso.";
                                                                    }
                                                                }
                                                            }
                                                            else 
                                                            {
                                                                Process p = new System.Diagnostics.Process();
                                                                ProcessStartInfo startInfo = new ProcessStartInfo();
                                                                p.StartInfo.FileName = internalSumatraPdfpath;
                                                                if (internaldebug == true)
                                                                { p.StartInfo.Arguments = "-exit-when-done -print-to \"" + printer + "\" " + a.FullName; }
                                                                else
                                                                {
                                                                    p.StartInfo.Arguments = "-silent -exit-when-done -print-to \"" + printer + "\" " + a.FullName;
                                                                }
                                                                p.Start();
                                                                p.WaitForExit();
                                                                erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " foi enviado para Booklet."));
                                                                Thread.Sleep(500);
                                                                //MessageBox.Show(BookletAudit + " " + a.Name + DateTime.Now.ToString() + " " + internalBookletSecsToWait.ToString());
                                                                bool FoundPrint = Myhelp.FoundPrintJob(BookletAudit, a.Name, DateTime.Now, internalBookletSecsToWait);
                                                                if (FoundPrint == true)
                                                                {
                                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " foi encontrado em " + AuditIdentifier));
                                                                    rp[4] = "Impresso.";
                                                                }
                                                                else
                                                                {
                                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " não foi encontrado em " + AuditIdentifier));
                                                                    rp[4] = "Impresso.";
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                else//standard print
                                                {
                                                    Process p = new System.Diagnostics.Process();
                                                    ProcessStartInfo startInfo = new ProcessStartInfo();
                                                    p.StartInfo.FileName = internalSumatraPdfpath;
                                                    p.StartInfo.Arguments = "-silent -exit-when-done -print-to \"" + printer + "\" " + a.FullName; ;
                                                    p.Start();
                                                    p.WaitForExit();
                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + "  foi enviado para " + printer));
                                                    Thread.Sleep(500);
                                                    rp[4] = "Impresso.";
                                                }

                                            }
                                            else { erros.Add(Myhelp.Formatmsg("A impressora " + printer + " não é valida.")); }
                                        }
                                        else { rp[4] = "Impresso."; }
                                    }
                                    else { erros.Add(Myhelp.Formatmsg(internalSumatraPdfpath + " não está disponivel.")); }

                                } 
                                #endregion
                                #region Caso Reset
                                if (rp[4].Contains("reset"))
                                {
                                    rp[4] = "Aguarda impressao.";
                                }
                                #endregion
                                #region Caso Apagar
                                if (rp[4].Contains("Apagar"))
                                {
                                    if (File.Exists(pathtoanalyze + @"/" + rp[6] + ".pdf"))
                                    {
                                        if (Myhelp.FileInUse(pathtoanalyze + @"/" + rp[6] + ".pdf") == false)
                                        {
                                            bool sucess = false;
                                            try { File.Delete(pathtoanalyze + @"/" + rp[6] + ".pdf"); sucess = true; }
                                            catch { sucess = false; }
                                            if (sucess == true)
                                            {
                                                try { File.Delete(pathtoanalyze + @"/" + rp[6] + ".txt"); }
                                                catch {  }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                            foreach (List<string> im in img)
                            {
                                #region Caso Armazenado
                                if (im[4].Contains("Armazenado."))
                                {
                                    if (im[8] != "default")
                                    {
                                        FileInfo a = new FileInfo(pathtoanalyze + @"\" + im[7] + ".pdf");
                                        FileInfo b = new FileInfo(pathtoanalyze + @"\" + im[7] + ".txt");
                                        if (a.CreationTime.AddHours(4) < DateTime.Now)
                                        {
                                            if (Myhelp.FileInUse(a.FullName) == false && Myhelp.FileInUse(b.FullName) == false)
                                            {
                                                try { File.Delete(a.FullName); File.Delete(b.FullName); }
                                                catch { erros.Add(Myhelp.Formatmsg("Não foi possível apagar " + a.FullName)); }//erro ao apagar ficheiros
                                            }
                                        }
                                    }
                                }
                                #endregion
                                #region Caso Aguarda armazenamento
                                if (im[4].Contains("Aguarda armazenamento."))
                                {
                                    if (im[3] != "default")
                                    {
                                        bool inerror = false;
                                        if (Directory.Exists(internalstoreimages  + @"\" + im[3]) == false)
                                        { Directory.CreateDirectory(internalstoreimages + @"\" + im[3]); Thread.Sleep(500); }
                                        try { File.Copy(pathtoanalyze + @"\" + im[7] + ".pdf", internalstoreimages + @"\" + im[3] + @"\" + im[5] + "." + im[1] + ".pdf", true); File.Copy(pathtoanalyze + @"\" + im[7] + ".txt", internalstoreimages + @"\" + im[3] + @"\" + im[5] + "." + im[1] + ".txt", true); }
                                        catch { inerror = true; erros.Add(Myhelp.Formatmsg("Não foi possível copiar " + pathtoanalyze + @"\" + im[7] + ".pdf")); }//mensagem de erro ao copiar
                                        if (inerror == false)
                                        { im[4] = "Armazenado."; }
                                    }
                                }
                                #endregion
                                #region Aguarda Relatorio
                                if (im[4].Contains("Aguarda Relatorio."))
                                {
                                    if (internalTempjoin == true)
                                    {
                                        if (im[6] == "default" && im[8] == "default")
                                        {
                                            foreach (List<string> rp in rep)
                                            {
                                                if (rp[4].Contains("Aguarda Imagem."))
                                                {
                                                    if (rp[7] == "default" && rp[8] == "default")
                                                    {
                                                        if (im[3] == "default" && rp[3] != "default")
                                                        { im[3] = rp[3]; }
                                                        List<string> newcomplete = new List<string>();
                                                        newcomplete.Add(rp[0]);
                                                        newcomplete.Add(rp[1]);
                                                        newcomplete.Add("C");
                                                        newcomplete.Add(rp[3]);
                                                        newcomplete.Add("Aguarda impressao.");
                                                        if (pathtoanalyze == internalpathECG)
                                                        { newcomplete.Add("ECG"); }
                                                        if (pathtoanalyze == internalpathECO)
                                                        { newcomplete.Add("ECO"); }
                                                        if (pathtoanalyze == internalpathMAPA)
                                                        { newcomplete.Add("MAPA"); }
                                                        if (pathtoanalyze == internalpathHOLTER)
                                                        { newcomplete.Add("HOLTER"); }
                                                        if (pathtoanalyze == internalpathPE)
                                                        { newcomplete.Add("PE"); }
                                                        newcomplete.Add(rp[6]);
                                                        newcomplete.Add(im[7]);
                                                        string targetfile = pathtoanalyze + @"\" + rp[5] + "C";
                                                        int x = 0;
                                                        if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                        {
                                                            while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                            { x++; }
                                                        }
                                                        string completefilename = targetfile + x.ToString() + ".pdf";
                                                        newcomplete.Add(rp[5] + "C" + x.ToString());
                                                        if (Myhelp.FileInUse(pathtoanalyze + @"\" + rp[6] + ".pdf") == false && Myhelp.FileInUse(pathtoanalyze + @"\" + im[7] + ".pdf") == false)
                                                        {
                                                            string[] files = new string[] { pathtoanalyze + @"\" + rp[6] + ".pdf", pathtoanalyze + @"\" + im[7] + ".pdf" };
                                                            bool sucess = Myhelp.MergePDF(completefilename, files);
                                                            if (sucess == true)
                                                            {
                                                                rp[8] = rp[5] + "C" + x.ToString();
                                                                rp[7] = im[7];
                                                                rp[4] = "Aguarda armazenamento.";
                                                                im[8] = rp[5] + "C" + x.ToString();
                                                                im[6] = rp[6];
                                                                im[4] = "Aguarda armazenamento.";
                                                                cpt.Add(newcomplete);
                                                            }
                                                        }
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                                #region Caso Impresso
                                if (im[4].Contains("Impresso."))
                                {
                                    im[4] = "Aguarda Relatorio.";
                                }
                                #endregion
                                #region Caso Aguarda impressao
                                if (im[4].Contains("Aguarda impressao."))
                                {
                                    bool toprint = internalTempIPrint;
                                    string printer = internalTempIPrinter;
                                    FileInfo a = new FileInfo(pathtoanalyze + @"\" + im[7] + ".pdf");

                                    if (File.Exists(internalSumatraPdfpath))
                                    {
                                        if (toprint == true)
                                        {
                                            if (Myhelp.printerisvalid(printer))
                                            {
                                                if (printer == internalBookletPrinter)//Bookletprint
                                                {
                                                    bool procrunning = Myhelp.isProcessRunning(internalBookletProcess);
                                                    if (procrunning == false)
                                                    {
                                                        if (Directory.Exists(internalBookletAudit))
                                                        {
                                                            string AuditIdentifier = "Print-in." + DateTime.Now.ToString("yyyyMMdd") + ".xml";
                                                            if (File.Exists(internalBookletAudit + @"\" + AuditIdentifier))
                                                            {
                                                                if (internaldebug == true)
                                                                {
                                                                    erros.Add(Myhelp.Formatmsg("Última impressão a :" + Myhelp.LastBookletPrintTime(internalBookletAudit).ToString() + " Com SecsTowait - " + Myhelp.LastBookletPrintTime(internalBookletAudit).AddSeconds(internalBookletSecsToWait).ToString() + " Now: " + DateTime.Now.ToString()));
                                                                }
                                                                if (Myhelp.LastBookletPrintTime(internalBookletAudit).AddSeconds(internalBookletSecsToWait) < DateTime.Now)
                                                                {
                                                                    Process p = new System.Diagnostics.Process();
                                                                    ProcessStartInfo startInfo = new ProcessStartInfo();
                                                                    p.StartInfo.FileName = internalSumatraPdfpath;
                                                                    if (internaldebug == true)
                                                                    { p.StartInfo.Arguments = "-exit-when-done -print-to \"" + printer + "\" " + a.FullName; }
                                                                    else
                                                                    {
                                                                        p.StartInfo.Arguments = "-silent -exit-when-done -print-to \"" + printer + "\" " + a.FullName;
                                                                    }
                                                                    p.Start();
                                                                    p.WaitForExit();
                                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " foi enviado para Booklet."));
                                                                    Thread.Sleep(500);
                                                                    //MessageBox.Show(BookletAudit + " " + a.Name + DateTime.Now.ToString() + " " + internalBookletSecsToWait.ToString());
                                                                    bool FoundPrint = Myhelp.FoundPrintJob(BookletAudit, a.Name, DateTime.Now, internalBookletSecsToWait);
                                                                    if (FoundPrint == true)
                                                                    {
                                                                        erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " foi encontrado em " + AuditIdentifier));
                                                                        im[4] = "Impresso.";
                                                                    }
                                                                    else
                                                                    {
                                                                        erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " não foi encontrado em " + AuditIdentifier));
                                                                        im[4] = "Impresso.";
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                Process p = new System.Diagnostics.Process();
                                                                ProcessStartInfo startInfo = new ProcessStartInfo();
                                                                p.StartInfo.FileName = internalSumatraPdfpath;
                                                                if (internaldebug == true)
                                                                { p.StartInfo.Arguments = "-exit-when-done -print-to \"" + printer + "\" " + a.FullName; }
                                                                else
                                                                {
                                                                    p.StartInfo.Arguments = "-silent -exit-when-done -print-to \"" + printer + "\" " + a.FullName;
                                                                }
                                                                p.Start();
                                                                p.WaitForExit();
                                                                erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " foi enviado para Booklet."));
                                                                Thread.Sleep(500);
                                                                //MessageBox.Show(BookletAudit + " " + a.Name + DateTime.Now.ToString() + " " + internalBookletSecsToWait.ToString());
                                                                bool FoundPrint = Myhelp.FoundPrintJob(BookletAudit, a.Name, DateTime.Now, internalBookletSecsToWait);
                                                                if (FoundPrint == true)
                                                                {
                                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " foi encontrado em " + AuditIdentifier));
                                                                    im[4] = "Impresso.";
                                                                }
                                                                else
                                                                {
                                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " não foi encontrado em " + AuditIdentifier));
                                                                    im[4] = "Impresso.";
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                else//standard print
                                                {
                                                    Process p = new System.Diagnostics.Process();
                                                    ProcessStartInfo startInfo = new ProcessStartInfo();
                                                    p.StartInfo.FileName = internalSumatraPdfpath;
                                                    p.StartInfo.Arguments = "-silent -exit-when-done -print-to \"" + printer + "\" " + a.FullName; ;
                                                    p.Start();
                                                    p.WaitForExit();
                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + "  foi enviado para " + printer));
                                                    Thread.Sleep(500);
                                                    im[4] = "Impresso.";
                                                }

                                            }
                                            else { erros.Add(Myhelp.Formatmsg("A impressora " + printer + " não é valida.")); }
                                        }
                                        else { im[4] = "Impresso."; }
                                    }
                                    else { erros.Add(Myhelp.Formatmsg(internalSumatraPdfpath + " não está disponivel.")); }

                                }
                                #endregion
                                #region Caso Reset
                                if (im[4].Contains("reset"))
                                {
                                    im[4] = "Aguarda impressao.";
                                }
                                #endregion
                                #region Caso Apagar
                                if (im[4].Contains("Apagar"))
                                {
                                    if (File.Exists(pathtoanalyze + @"/" + im[7] + ".pdf"))
                                    {
                                        if (Myhelp.FileInUse(pathtoanalyze + @"/" + im[7] + ".pdf") == false)
                                        {
                                            bool sucess = false;
                                            try { File.Delete(pathtoanalyze + @"/" + im[7] + ".pdf"); sucess = true; }
                                            catch { sucess = false; }
                                            if (sucess == true)
                                            {
                                                try { File.Delete(pathtoanalyze + @"/" + im[7] + ".txt"); }
                                                catch { }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                            foreach (List<string> cp in cpt)
                            {
                                #region Caso Armazenado
                                if (cp[4].Contains("Armazenado."))
                                {
                                    if (cp[8] != "default")
                                    {
                                        FileInfo a = new FileInfo(pathtoanalyze + @"\" + cp[8] + ".pdf");
                                        FileInfo b = new FileInfo(pathtoanalyze + @"\" + cp[8] + ".txt");
                                        if (a.CreationTime.AddHours(4) < DateTime.Now)
                                        {
                                            if (Myhelp.FileInUse(a.FullName) == false && Myhelp.FileInUse(b.FullName) == false)
                                            {
                                                try { File.Delete(a.FullName); File.Delete(b.FullName); }
                                                catch { erros.Add(Myhelp.Formatmsg("Não foi possível apagar " + a.FullName)); }//erro ao apagar ficheiros
                                            }
                                        }
                                    }
                                }
                                #endregion
                                #region Caso Aguarda armazenamento
                                if (cp[4].Contains("Aguarda armazenamento."))
                                {
                                    if (cp[3] != "default")
                                    {
                                        bool inerror = false;
                                        if (Directory.Exists(internalstorecomplete + @"\" + cp[3]) == false)
                                        { Directory.CreateDirectory(internalstorecomplete + @"\" + cp[3]); Thread.Sleep(500); }
                                        try { File.Copy(pathtoanalyze + @"\" + cp[8] + ".pdf", internalstorecomplete + @"\" + cp[3] + @"\" + cp[5] + "." + cp[1] + ".pdf", true); File.Copy(pathtoanalyze + @"\" + cp[8] + ".txt", internalstorecomplete + @"\" + cp[3] + @"\" + cp[5] + "." + cp[1] + ".txt", true); }
                                        catch { inerror = true; erros.Add(Myhelp.Formatmsg("Não foi possível copiar " + pathtoanalyze + @"\" + cp[8] + ".pdf")); }//mensagem de erro ao copiar
                                        if (inerror == false)
                                        { cp[4] = "Armazenado."; }
                                    }
                                }
                                #endregion
                                #region Caso Impresso
                                if (cp[4].Contains("Impresso."))
                                {
                                    cp[4] = "Aguarda armazenamento.";
                                }
                                #endregion
                                #region Caso Aguarda impressao
                                if (cp[4].Contains("Aguarda impressao."))
                                {
                                    bool toprint = internalTempCPrint;
                                    string printer = internalTempCPrinter;
                                    FileInfo a = new FileInfo(pathtoanalyze + @"\" + cp[8] + ".pdf");

                                    if (File.Exists(internalSumatraPdfpath))
                                    {
                                        if (toprint == true)
                                        {
                                            if (Myhelp.printerisvalid(printer))
                                            {
                                                if (printer == internalBookletPrinter)//Bookletprint
                                                {
                                                    bool procrunning = Myhelp.isProcessRunning(internalBookletProcess);
                                                    if (procrunning == false)
                                                    {
                                                        if (Directory.Exists(internalBookletAudit))
                                                        {
                                                            string AuditIdentifier = "Print-in." + DateTime.Now.ToString("yyyyMMdd") + ".xml";
                                                            if (File.Exists(internalBookletAudit + @"\" + AuditIdentifier))
                                                            {
                                                                if (internaldebug == true)
                                                                {
                                                                    erros.Add(Myhelp.Formatmsg("Última impressão a :" + Myhelp.LastBookletPrintTime(internalBookletAudit).ToString() + " Com SecsTowait - " + Myhelp.LastBookletPrintTime(internalBookletAudit).AddSeconds(internalBookletSecsToWait).ToString() + " Now: " + DateTime.Now.ToString()));
                                                                }
                                                                if (Myhelp.LastBookletPrintTime(internalBookletAudit).AddSeconds(internalBookletSecsToWait) < DateTime.Now)
                                                                {
                                                                    Process p = new System.Diagnostics.Process();
                                                                    ProcessStartInfo startInfo = new ProcessStartInfo();
                                                                    p.StartInfo.FileName = internalSumatraPdfpath;
                                                                    if (internaldebug == true)
                                                                    { p.StartInfo.Arguments = "-exit-when-done -print-to \"" + printer + "\" " + a.FullName; }
                                                                    else
                                                                    {
                                                                        p.StartInfo.Arguments = "-silent -exit-when-done -print-to \"" + printer + "\" " + a.FullName;
                                                                    }
                                                                    p.Start();
                                                                    p.WaitForExit();
                                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " foi enviado para Booklet."));
                                                                    Thread.Sleep(500);
                                                                    //MessageBox.Show(BookletAudit + " " + a.Name + DateTime.Now.ToString() + " " + internalBookletSecsToWait.ToString());
                                                                    bool FoundPrint = Myhelp.FoundPrintJob(BookletAudit, a.Name, DateTime.Now, internalBookletSecsToWait);
                                                                    if (FoundPrint == true)
                                                                    {
                                                                        erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " foi encontrado em " + AuditIdentifier));
                                                                        cp[4] = "Impresso.";
                                                                    }
                                                                    else
                                                                    {
                                                                        erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " não foi encontrado em " + AuditIdentifier));
                                                                        cp[4] = "Impresso.";
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                Process p = new System.Diagnostics.Process();
                                                                ProcessStartInfo startInfo = new ProcessStartInfo();
                                                                p.StartInfo.FileName = internalSumatraPdfpath;
                                                                if (internaldebug == true)
                                                                { p.StartInfo.Arguments = "-exit-when-done -print-to \"" + printer + "\" " + a.FullName; }
                                                                else
                                                                {
                                                                    p.StartInfo.Arguments = "-silent -exit-when-done -print-to \"" + printer + "\" " + a.FullName;
                                                                }
                                                                p.Start();
                                                                p.WaitForExit();
                                                                erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " foi enviado para Booklet."));
                                                                Thread.Sleep(500);
                                                                //MessageBox.Show(BookletAudit + " " + a.Name + DateTime.Now.ToString() + " " + internalBookletSecsToWait.ToString());
                                                                bool FoundPrint = Myhelp.FoundPrintJob(BookletAudit, a.Name, DateTime.Now, internalBookletSecsToWait);
                                                                if (FoundPrint == true)
                                                                {
                                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " foi encontrado em " + AuditIdentifier));
                                                                    cp[4] = "Impresso.";
                                                                }
                                                                else
                                                                {
                                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + " não foi encontrado em " + AuditIdentifier));
                                                                    cp[4] = "Impresso.";
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                else//standard print
                                                {
                                                    Process p = new System.Diagnostics.Process();
                                                    ProcessStartInfo startInfo = new ProcessStartInfo();
                                                    p.StartInfo.FileName = internalSumatraPdfpath;
                                                    p.StartInfo.Arguments = "-silent -exit-when-done -print-to \"" + printer + "\" " + a.FullName; ;
                                                    p.Start();
                                                    p.WaitForExit();
                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name + "  foi enviado para " + printer));
                                                    Thread.Sleep(500);
                                                    cp[4] = "Impresso.";
                                                }

                                            }
                                            else { erros.Add(Myhelp.Formatmsg("A impressora " + printer + " não é valida.")); }
                                        }
                                        else { cp[4] = "Impresso."; }
                                    }
                                    else { erros.Add(Myhelp.Formatmsg(internalSumatraPdfpath + " não está disponivel.")); }

                                }
                                #endregion
                                #region Caso Reset
                                if (cp[4].Contains("reset"))
                                {
                                    cp[4] = "Aguarda impressao.";
                                }
                                #endregion
                                #region Caso Apagar
                                if (cp[4].Contains("Apagar"))
                                {
                                    if (File.Exists(pathtoanalyze + @"/" + cp[8] + ".pdf"))
                                    {
                                        if (Myhelp.FileInUse(pathtoanalyze + @"/" + cp[8] + ".pdf") == false)
                                        {
                                            bool sucess = false;
                                            try { File.Delete(pathtoanalyze + @"/" + cp[8] + ".pdf"); sucess = true; }
                                            catch { sucess = false; }
                                            if (sucess == true)
                                            {
                                                try { File.Delete(pathtoanalyze + @"/" + cp[8] + ".txt"); }
                                                catch { }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                            Thread.Sleep(100);
                            #region Update Reps
                            foreach (List<string> rp in rep)
                            {
                                Thread.Sleep(200);
                                if (rp.Count >= 8)
                                {
                                    string repfile = pathtoanalyze + @"\" + rp[6] + ".txt";
                                    if (File.Exists(repfile))
                                    {
                                        if (Myhelp.FileInUse(repfile) == false)
                                        {
                                            List<string> currentcontent = new List<string>();
                                            try
                                            {
                                                currentcontent = File.ReadAllLines(repfile).ToList();
                                                
                                            }
                                            catch { }
                                            if (currentcontent.Count >= 8)
                                            {
                                                if (currentcontent.SequenceEqual(rp)== false)
                                                {
                                                    
                                                    rp[1] = currentcontent[1];
                                                    try { File.WriteAllLines(repfile, rp.ToArray());  }
                                                    catch { erros.Add(Myhelp.Formatmsg("Não foi possível atualizar " + repfile)); }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        try { File.WriteAllLines(repfile, rp.ToArray());  }
                                        catch { erros.Add(Myhelp.Formatmsg("Não foi possível atualizar " + repfile)); }
                                    }
                                }
                            } 
                            #endregion
                            #region Update Imgs
                            foreach (List<string> im in img)
                            {
                                Thread.Sleep(200);
                                if (im.Count >= 8)
                                {
                                    string repfile = pathtoanalyze + @"\" + im[7] + ".txt";
                                    if (File.Exists(repfile))
                                    {
                                        if (Myhelp.FileInUse(repfile) == false)
                                        {
                                            List<string> currentcontent = new List<string>();
                                            try
                                            {
                                                currentcontent = File.ReadAllLines(repfile).ToList();
                                            }
                                            catch { }
                                            if (currentcontent.Count >= 8)
                                            {
                                                if (currentcontent.SequenceEqual(im) == false)
                                                {
                                                    
                                                    im[1] = currentcontent[1];
                                                    try { File.WriteAllLines(repfile, im.ToArray());  }
                                                    catch { erros.Add(Myhelp.Formatmsg("Não foi possível atualizar " + repfile)); }
                                                }
                                            }
                                            
                                        }
                                    }
                                    else
                                    {
                                        try { File.WriteAllLines(repfile, im.ToArray());  }
                                        catch { erros.Add(Myhelp.Formatmsg("Não foi possível atualizar " + repfile)); }
                                    }
                                }
                            } 
                            #endregion
                            #region Update Cpts
                            foreach (List<string> cp in cpt)
                            {
                                Thread.Sleep(200);
                                if (cp.Count >= 8)
                                {
                                    string repfile = pathtoanalyze + @"\" + cp[8] + ".txt";
                                    if (File.Exists(repfile))
                                    {
                                        if (Myhelp.FileInUse(repfile) == false)
                                        {
                                            List<string> currentcontent = new List<string>();
                                            try
                                            {
                                                currentcontent = File.ReadAllLines(repfile).ToList();
                                            }
                                            catch { }
                                            if (currentcontent.Count >= 8)
                                            {
                                                if (currentcontent.SequenceEqual(cp) == false)
                                                {
                                                    
                                                    cp[1] = currentcontent[1];
                                                    try { File.WriteAllLines(repfile, cp.ToArray());  }
                                                    catch { erros.Add(Myhelp.Formatmsg("Não foi possível atualizar " + repfile)); }
                                                }
                                            }
                                            
                                        }
                                    }
                                    else
                                    {
                                        try { File.WriteAllLines(repfile, cp.ToArray());  }
                                        catch { erros.Add(Myhelp.Formatmsg("Não foi possível atualizar " + repfile)); }
                                    }
                                }
                            } 
                            #endregion

                        }

                    }
                    else { erros.Add(Myhelp.Formatmsg(" A directoria " + pathtoanalyze + " não existe ou não está disponível.")); }
                    #region OldFolders
                    /*
                    //MessageBox.Show(pathtoanalyze + Environment.NewLine + internalTempRPrinter + Environment.NewLine + internalTempRPrint.ToString() + internalTempIPrinter + internalTempCPrinter);
                    if (Directory.Exists(pathtoanalyze))
                    {
                        DirectoryInfo di = new DirectoryInfo(pathtoanalyze);
                        List<FileInfo> PathFiles = di.GetFiles("*.txt", SearchOption.TopDirectoryOnly).ToList();
                        if (PathFiles.Count > 0)
                        {
                            List<FileInfo> orderedList = PathFiles.OrderBy(x => x.CreationTime).ToList();
                            List<string> ProcessedFiles = new List<string>();
                            foreach (FileInfo a in orderedList)
                            {
                                string infoFile = a.FullName; string pdfFile = a.FullName.Replace(".txt", ".pdf");
                                if (File.Exists(infoFile) && File.Exists(pdfFile) && ProcessedFiles.Contains(infoFile.Replace(".txt", "")) == false)
                                {
                                    Thread.Sleep(100);
                                    List<string> FileProp = new List<string>();
                                    List<string> oldFileProp = new List<string>();
                                    oldFileProp.AddRange(FileProp);
                                    FileProp = File.ReadAllLines(infoFile).ToList();
                                    List<string> RepFileProp = new List<string>();
                                    List<string> ImgFileProp = new List<string>();
                                    List<string> CompFileProp = new List<string>();
                                    if (FileProp.Count >= 8)
                                    {

                                        #region Mudar tipo de ficheiro (pasta)
                                        string BelongingDir = "";
                                        if (FileProp[5] == "ECG")
                                        { BelongingDir = internalpathECG; }
                                        if (FileProp[5] == "ECO")
                                        { BelongingDir = internalpathECO; }
                                        if (FileProp[5] == "HOLTER")
                                        { BelongingDir = internalpathHOLTER; }
                                        if (FileProp[5] == "MAPA")
                                        { BelongingDir = internalpathMAPA; }
                                        if (FileProp[5] == "PE")
                                        { BelongingDir = internalpathPE; }
                                        if (FileProp[5] == "UC")
                                        { BelongingDir = internalpathUC; }
                                        // mudar o tipo de exame
                                        if (pathtoanalyze != BelongingDir)
                                        {
                                            if (Myhelp.FileInUse(pdfFile) == false)
                                            {
                                                string targetfile = BelongingDir + @"\" + FileProp[5] + FileProp[2];
                                                int x = 0;
                                                if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                {
                                                    while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                    { x++; }
                                                }
                                                if (FileProp[2] == "R")
                                                {
                                                    FileProp[6] = FileProp[5] + FileProp[2] + x.ToString();
                                                }
                                                if (FileProp[2] == "I")
                                                {
                                                    FileProp[7] = FileProp[5] + FileProp[2] + x.ToString();
                                                }
                                                if (FileProp[2] == "UC")
                                                {
                                                    FileProp[6] = FileProp[5] + FileProp[2] + x.ToString();
                                                }
                                                if (Directory.Exists(BelongingDir))
                                                {

                                                    try
                                                    {
                                                        File.Move(pdfFile, targetfile + x.ToString() + ".pdf");
                                                        File.WriteAllLines(targetfile + x.ToString() + ".txt", FileProp.ToArray());
                                                    }
                                                    catch { erros.Add(Myhelp.Formatmsg("Não foi possível criar o ficheiro: " + targetfile + x.ToString() + ".txt")); }
                                                    erros.Add(Myhelp.Formatmsg(a.FullName.Replace(".txt", ".pdf") + " movido para " + BelongingDir + " com o nome:" + targetfile + x.ToString() + ".pdf"));
                                                }
                                            }
                                            continue;
                                        #endregion
                                        }//mudar de exame
                                        else
                                        {
                                            #region Procurar Relatorios
                                            if (FileProp[2] == "R")
                                            { RepFileProp.Clear(); RepFileProp.AddRange(FileProp); }
                                            if (FileProp[2] == "I")
                                            { ImgFileProp.Clear(); ImgFileProp.AddRange(FileProp); }
                                            if (FileProp[2] == "C")
                                            { CompFileProp.Clear(); CompFileProp.AddRange(FileProp); }
                                            if (FileProp[6] == "default")
                                            {
                                                foreach (FileInfo b in orderedList)
                                                {
                                                    List<string> FilePropPartner = new List<string>();
                                                    FilePropPartner = File.ReadAllLines(b.FullName).ToList();
                                                    if (FilePropPartner[2] == "R" && FilePropPartner[1] == FileProp[1] && FileProp[2] == "I" && FileProp[6] == "default" && ProcessedFiles.Contains(b.Name.Replace(".txt", "")) == false)
                                                    {
                                                        if (FilePropPartner[4] == "Aguarda impressao." || FilePropPartner[4] == "Impresso." || FilePropPartner[4] == "Aguarda Imagem." || FilePropPartner[4] == "Aguarda Relatorio.")
                                                        {
                                                            FileProp[6] = FilePropPartner[6]; RepFileProp.Clear(); RepFileProp.AddRange(FilePropPartner); ProcessedFiles.Add(FileProp[6]); break;
                                                        }

                                                    }
                                                }
                                            }
                                            else
                                            {
                                                List<string> FilePropPartner = new List<string>();
                                                if (File.Exists(pathtoanalyze + @"\" + FileProp[6] + ".txt") && ProcessedFiles.Contains(FileProp[6]) == false)
                                                {
                                                    try
                                                    {
                                                        FilePropPartner = File.ReadAllLines(pathtoanalyze + @"\" + FileProp[6] + ".txt").ToList();
                                                        if (FilePropPartner.Count >= 8)
                                                        {
                                                            RepFileProp.Clear(); RepFileProp.AddRange(FilePropPartner); ProcessedFiles.Add(FileProp[6]);
                                                        }
                                                    }
                                                    catch { }
                                                }
                                            }
                                            #endregion
                                            #region Procurar Imagens
                                            if (FileProp[7] == "false" || FileProp[7] == "default")
                                            {
                                                if (FileProp[7] == "default")
                                                {
                                                    if (internalTempjoin == false)
                                                    {
                                                        FileProp[7] = "false";
                                                    }
                                                    else
                                                    {
                                                        foreach (FileInfo b in orderedList)
                                                        {
                                                            List<string> FilePropPartner = new List<string>();
                                                            FilePropPartner = File.ReadAllLines(b.FullName).ToList();
                                                            if (FilePropPartner[2] == "I" && FilePropPartner[1] == FileProp[1] && FileProp[2] == "R" && FileProp[7] == "default" && ProcessedFiles.Contains(FileProp[7]) == false)
                                                            {
                                                                if (FilePropPartner[4] == "Aguarda impressao." || FilePropPartner[4] == "Impresso." || FilePropPartner[4] == "Aguarda Imagem." || FilePropPartner[4] == "Aguarda Relatorio." && ProcessedFiles.Contains(FileProp[7]) == false)
                                                                {
                                                                    FileProp[7] = FilePropPartner[7]; ImgFileProp.Clear(); ImgFileProp.AddRange(FilePropPartner); ProcessedFiles.Add(FileProp[7]); break;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                List<string> FilePropPartner = new List<string>();
                                                if (File.Exists(pathtoanalyze + @"\" + FileProp[7] + ".txt") && ProcessedFiles.Contains(FileProp[7]) == false)
                                                {
                                                    try
                                                    {
                                                        FilePropPartner = File.ReadAllLines(pathtoanalyze + @"\" + FileProp[7] + ".txt").ToList();
                                                        if (FilePropPartner.Count >= 8)
                                                        {
                                                            ImgFileProp.Clear(); ImgFileProp.AddRange(FilePropPartner); ProcessedFiles.Add(FileProp[7]);
                                                        }
                                                    }
                                                    catch { }
                                                }
                                            }
                                            #endregion
                                            #region Procurar Completos
                                            if (FileProp[8] == "default" || FileProp[8] == "missing")
                                            {
                                                foreach (FileInfo b in orderedList)
                                                {
                                                    List<string> FilePropPartner = new List<string>();
                                                    FilePropPartner = File.ReadAllLines(b.FullName).ToList();
                                                    if (FilePropPartner[2] == "C" && FilePropPartner[1] == FileProp[1] && FileProp[6] == FilePropPartner[6] && FileProp[7] == FilePropPartner[7] && ProcessedFiles.Contains(FileProp[8]) == false)
                                                    {

                                                        if (FilePropPartner[4] == "Aguarda impressao." || FilePropPartner[4] == "Impresso." || FilePropPartner[4] == "Aguarda Imagem." || FilePropPartner[4] == "Aguarda Relatorio.")
                                                        {
                                                            FileProp[8] = FilePropPartner[8]; CompFileProp.Clear(); CompFileProp.AddRange(FilePropPartner); ProcessedFiles.Add(FileProp[8]); break;
                                                        }

                                                    }
                                                }
                                            }
                                            else
                                            {
                                                List<string> FilePropPartner = new List<string>();
                                                if (File.Exists(pathtoanalyze + @"\" + FileProp[8] + ".txt"))
                                                {
                                                    try
                                                    {
                                                        FilePropPartner = File.ReadAllLines(pathtoanalyze + @"\" + FileProp[8] + ".txt").ToList();
                                                        if (FilePropPartner.Count >= 8)
                                                        {
                                                            if (FileProp[6] == FilePropPartner[6] && FileProp[7] == FilePropPartner[7])
                                                            {
                                                                CompFileProp.Clear(); CompFileProp.AddRange(FilePropPartner); ProcessedFiles.Add(FileProp[8]);
                                                            }
                                                        }
                                                    }
                                                    catch { }
                                                }
                                            }
                                            #endregion

                                            var areEquivalent = (FileProp.Count() == oldFileProp.Count()) && !FileProp.Except(oldFileProp).Any();
                                            if (areEquivalent == false)
                                            {

                                                File.WriteAllLines(infoFile, FileProp.ToArray());
                                                oldFileProp = FileProp;

                                            }
                                            // depois de atualizar o identificador para este ficheiro
                                            if (FileProp[4].Contains("Armazenado."))
                                            #region Armazenado
                                            {
                                                bool delete = false;
                                                if (RepFileProp.Count > 0 && ImgFileProp.Count > 0 && CompFileProp.Count > 0)
                                                {
                                                    if (RepFileProp[4].Contains("Armazenado.") && ImgFileProp[4].Contains("Armazenado.") && CompFileProp[4].Contains("Armazenado."))
                                                    {
                                                        if (a.CreationTime.AddHours(8) < DateTime.Now)
                                                        {
                                                            delete = true;
                                                        }
                                                    }
                                                    else { delete = false; }
                                                }
                                                else { delete = false; }
                                                if (delete == true)
                                                {
                                                    string RepFilePath = pathtoanalyze + @"\" + FileProp[6] + ".pdf";
                                                    string ImgFilePath = pathtoanalyze + @"\" + FileProp[7] + ".pdf";
                                                    string CompFilePath = pathtoanalyze + @"\" + FileProp[8] + ".pdf";
                                                    if (Myhelp.FileInUse(RepFilePath) == false && Myhelp.FileInUse(ImgFilePath) == false && Myhelp.FileInUse(CompFilePath) == false)
                                                    {
                                                        try
                                                        {
                                                            Thread.Sleep(500);
                                                            if (File.Exists(RepFilePath))
                                                            {
                                                                File.Delete(RepFilePath);
                                                            }
                                                            Thread.Sleep(500);
                                                            if (File.Exists(RepFilePath.Replace(".pdf", ".txt")))
                                                            {
                                                                File.Delete(RepFilePath.Replace(".pdf", ".txt"));
                                                            }
                                                            Thread.Sleep(500);
                                                            if (File.Exists(ImgFilePath))
                                                            {
                                                                File.Delete(ImgFilePath);
                                                            }
                                                            Thread.Sleep(500);
                                                            if (File.Exists(ImgFilePath.Replace(".pdf", ".txt")))
                                                            {
                                                                File.Delete(ImgFilePath.Replace(".pdf", ".txt"));
                                                            }
                                                            Thread.Sleep(500);
                                                            if (File.Exists(CompFilePath))
                                                            {
                                                                File.Delete(CompFilePath);
                                                            }
                                                            Thread.Sleep(500);
                                                            if (File.Exists(CompFilePath.Replace(".pdf", ".txt")))
                                                            {
                                                                File.Delete(CompFilePath.Replace(".pdf", ".txt"));
                                                            }
                                                            Thread.Sleep(500);
                                                        }
                                                        catch { }
                                                    }
                                                }
                                                if (FileProp[7] == "false")
                                                {
                                                    delete = false;
                                                    if (RepFileProp.Count > 0 && CompFileProp.Count > 0)
                                                    {
                                                        if (RepFileProp[4].Contains("Armazenado.") && CompFileProp[4].Contains("Armazenado."))
                                                        {
                                                            if (a.CreationTime.AddHours(8) < DateTime.Now)
                                                            {
                                                                delete = true;
                                                            }
                                                        }
                                                        else { delete = false; }
                                                    }
                                                    else { delete = false; }
                                                    if (delete == true)
                                                    {
                                                        string RepFilePath = pathtoanalyze + @"\" + FileProp[6] + ".pdf";

                                                        string CompFilePath = pathtoanalyze + @"\" + FileProp[8] + ".pdf";
                                                        if (Myhelp.FileInUse(RepFilePath) == false && Myhelp.FileInUse(CompFilePath) == false)
                                                        {
                                                            try
                                                            {
                                                                Thread.Sleep(500);
                                                                if (File.Exists(RepFilePath))
                                                                {
                                                                    File.Delete(RepFilePath);
                                                                }
                                                                Thread.Sleep(500);
                                                                if (File.Exists(RepFilePath.Replace(".pdf", ".txt")))
                                                                {
                                                                    File.Delete(RepFilePath.Replace(".pdf", ".txt"));
                                                                }
                                                                Thread.Sleep(500);

                                                                if (File.Exists(CompFilePath))
                                                                {
                                                                    File.Delete(CompFilePath);
                                                                }
                                                                Thread.Sleep(500);
                                                                if (File.Exists(CompFilePath.Replace(".pdf", ".txt")))
                                                                {
                                                                    File.Delete(CompFilePath.Replace(".pdf", ".txt"));
                                                                }
                                                                Thread.Sleep(500);
                                                            }
                                                            catch { }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                            if (FileProp[4] == "Aguarda armazenamento.")
                                            #region Aguarda armazenamento
                                            {
                                                bool storage = false;
                                                if (RepFileProp.Count > 0 && ImgFileProp.Count > 0 && CompFileProp.Count > 0)
                                                {
                                                    if (RepFileProp[4] == "Aguarda armazenamento." && ImgFileProp[4] == "Aguarda armazenamento." && CompFileProp[4] == "Aguarda armazenamento.")
                                                    {
                                                        if (Directory.Exists(internalstorereports) && Directory.Exists(internalstoreimages) && Directory.Exists(internalstorecomplete))
                                                        { storage = true; }
                                                        else { storage = false; }
                                                    }
                                                    else { storage = false; }
                                                }
                                                else { storage = false; }

                                                if (storage == true)
                                                #region Storage
                                                {
                                                    string RepFilePath = pathtoanalyze + @"\" + FileProp[6] + ".pdf";
                                                    string ImgFilePath = pathtoanalyze + @"\" + FileProp[7] + ".pdf";
                                                    string CompFilePath = pathtoanalyze + @"\" + FileProp[8] + ".pdf";
                                                    if (Directory.Exists(internalstorereports + @"\" + FileProp[3]) == false)
                                                    { Directory.CreateDirectory(internalstorereports + @"\" + FileProp[3]); }
                                                    if (Directory.Exists(internalstoreimages + @"\" + FileProp[3]) == false)
                                                    { Directory.CreateDirectory(internalstoreimages + @"\" + FileProp[3]); }
                                                    if (Directory.Exists(internalstorecomplete + @"\" + FileProp[3]) == false)
                                                    { Directory.CreateDirectory(internalstorecomplete + @"\" + FileProp[3]); }
                                                    if (Myhelp.FileInUse(RepFilePath) == false && Myhelp.FileInUse(ImgFilePath) == false && Myhelp.FileInUse(CompFilePath) == false)
                                                    {
                                                        string targetfile = internalstorereports + @"\" + FileProp[3] + @"\" + FileProp[5] + ".R." + RepFileProp[1];
                                                        int x = 0;
                                                        if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                        {
                                                            while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                            { x++; }
                                                        }
                                                        string completefilename = targetfile + x.ToString() + ".pdf";
                                                        try
                                                        {
                                                            File.Copy(RepFilePath, completefilename);
                                                            RepFileProp[4] = "Armazenado.";
                                                            File.WriteAllLines(RepFilePath.Replace(".pdf", ".txt"), RepFileProp.ToArray());
                                                            File.WriteAllLines(completefilename.Replace(".pdf", ".txt"), RepFileProp.ToArray());
                                                        }
                                                        catch { }
                                                        targetfile = internalstoreimages + @"\" + FileProp[3] + @"\" + FileProp[5] + ".I." + ImgFileProp[1];
                                                        x = 0;
                                                        if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                        {
                                                            while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                            { x++; }
                                                        }
                                                        completefilename = targetfile + x.ToString() + ".pdf";
                                                        try
                                                        {
                                                            File.Copy(ImgFilePath, completefilename);
                                                            ImgFileProp[4] = "Armazenado.";
                                                            File.WriteAllLines(ImgFilePath.Replace(".pdf", ".txt"), ImgFileProp.ToArray());
                                                            File.WriteAllLines(completefilename.Replace(".pdf", ".txt"), ImgFileProp.ToArray());
                                                        }
                                                        catch { }
                                                        targetfile = internalstorecomplete + @"\" + FileProp[3] + @"\" + FileProp[5] + ".C." + CompFileProp[1];
                                                        x = 0;
                                                        if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                        {
                                                            while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                            { x++; }
                                                        }
                                                        completefilename = targetfile + x.ToString() + ".pdf";
                                                        try
                                                        {
                                                            File.Copy(CompFilePath, completefilename);
                                                            CompFileProp[4] = "Armazenado.";
                                                            File.WriteAllLines(CompFilePath.Replace(".pdf", ".txt"), CompFileProp.ToArray());
                                                            File.WriteAllLines(completefilename.Replace(".pdf", ".txt"), CompFileProp.ToArray());
                                                        }
                                                        catch { }
                                                    }
                                                }
                                                #endregion
                                                if (FileProp[7] == "false")
                                                {
                                                    if (FileProp[8] == "default")//criar complete
                                                    #region Criar Complete
                                                    {

                                                        CompFileProp.AddRange(FileProp);
                                                        string RepFilePath = pathtoanalyze + @"\" + FileProp[6] + ".pdf";
                                                        string targetfile = pathtoanalyze + @"\" + FileProp[5] + "C";
                                                        int x = 0;
                                                        if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                        {
                                                            while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                            { x++; }
                                                        }
                                                        string completefilename = targetfile + x.ToString() + ".pdf";

                                                        FileProp[8] = FileProp[5] + "C" + x.ToString();

                                                        CompFileProp[8] = FileProp[5] + "C" + x.ToString();
                                                        CompFileProp[4] = "Aguarda impressao.";
                                                        CompFileProp[2] = "C";
                                                        File.Copy(RepFilePath, completefilename);
                                                        File.WriteAllLines(RepFilePath.Replace(".pdf", ".txt"), FileProp.ToArray());
                                                        File.WriteAllLines(completefilename.Replace(".pdf", ".txt"), CompFileProp.ToArray());
                                                    }
                                                    #endregion
                                                    else
                                                    {

                                                        if (RepFileProp.Count > 0 && CompFileProp.Count > 0)
                                                        {

                                                            if (RepFileProp[4] == "Aguarda armazenamento." && CompFileProp[4] == "Aguarda armazenamento.")
                                                            {

                                                                if (Directory.Exists(internalstorereports) && Directory.Exists(internalstorecomplete))
                                                                { storage = true; }
                                                                else { storage = false; }
                                                            }
                                                            else { storage = false; }
                                                        }
                                                        else { storage = false; }

                                                        if (storage == true)
                                                        {

                                                            string RepFilePath = pathtoanalyze + @"\" + FileProp[6] + ".pdf";
                                                            string CompFilePath = pathtoanalyze + @"\" + FileProp[8] + ".pdf";
                                                            if (Directory.Exists(internalstorereports + @"\" + FileProp[3]) == false)
                                                            { Directory.CreateDirectory(internalstorereports + @"\" + FileProp[3]); }
                                                            if (Directory.Exists(internalstorecomplete + @"\" + FileProp[3]) == false)
                                                            { Directory.CreateDirectory(internalstorecomplete + @"\" + FileProp[3]); }
                                                            if (Myhelp.FileInUse(RepFilePath) == false && Myhelp.FileInUse(CompFilePath) == false)
                                                            {
                                                                string targetfile = internalstorereports + @"\" + FileProp[3] + @"\" + FileProp[5] + ".R." + RepFileProp[1];
                                                                int x = 0;
                                                                if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                                {
                                                                    while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                                    { x++; }
                                                                }
                                                                string completefilename = targetfile + x.ToString() + ".pdf";
                                                                try
                                                                {
                                                                    File.Copy(RepFilePath, completefilename);
                                                                    RepFileProp[4] = "Armazenado.";
                                                                    File.WriteAllLines(RepFilePath.Replace(".pdf", ".txt"), RepFileProp.ToArray());
                                                                    File.WriteAllLines(completefilename.Replace(".pdf", ".txt"), RepFileProp.ToArray());
                                                                }
                                                                catch { }
                                                                targetfile = internalstorecomplete + @"\" + FileProp[3] + @"\" + FileProp[5] + ".C." + CompFileProp[1];
                                                                x = 0;
                                                                if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                                {
                                                                    while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                                    { x++; }
                                                                }
                                                                completefilename = targetfile + x.ToString() + ".pdf";
                                                                try
                                                                {
                                                                    File.Copy(CompFilePath, completefilename);
                                                                    CompFileProp[4] = "Armazenado.";
                                                                    File.WriteAllLines(CompFilePath.Replace(".pdf", ".txt"), CompFileProp.ToArray());
                                                                    File.WriteAllLines(completefilename.Replace(".pdf", ".txt"), CompFileProp.ToArray());
                                                                }
                                                                catch { }

                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                            if (FileProp[4] == "Aguarda Imagem." || FileProp[4] == "Aguarda Relatorio.")
                                            #region Caso Juntar
                                            {
                                                if (internalTempjoin == true)
                                                {
                                                    if (RepFileProp.Count > 0 && ImgFileProp.Count > 0 && CompFileProp.Count == 0)
                                                    {
                                                        string RepFilePath = pathtoanalyze + @"\" + FileProp[6] + ".pdf";
                                                        string ImgFilePath = pathtoanalyze + @"\" + FileProp[7] + ".pdf";
                                                        if (ImgFileProp[4] == "Aguarda Relatorio." && RepFileProp[4] == "Aguarda Imagem.")
                                                        {
                                                            if (Myhelp.FileInUse(RepFilePath) == false && Myhelp.FileInUse(ImgFilePath) == false)
                                                            {
                                                                string targetfile = pathtoanalyze + @"\" + FileProp[5] + "C";
                                                                int x = 0;
                                                                if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                                {
                                                                    while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                                    { x++; }
                                                                }
                                                                string completefilename = targetfile + x.ToString() + ".pdf";
                                                                string[] files = new string[] { RepFilePath, ImgFilePath };
                                                                bool sucess = Myhelp.MergePDF(completefilename, files);
                                                                if (sucess == true)
                                                                {
                                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + completefilename + " foi criado com sucesso."));
                                                                    try
                                                                    {

                                                                        CompFileProp.AddRange(FileProp);
                                                                        CompFileProp[2] = "C";
                                                                        CompFileProp[4] = "Aguarda impressao.";
                                                                        if (CompFileProp[3] == "default" && CompFileProp[3] != RepFileProp[3])
                                                                        { CompFileProp[3] = RepFileProp[3]; }
                                                                        CompFileProp[8] = FileProp[5] + "C" + x.ToString();
                                                                        FileProp[4] = "Aguarda armazenamento.";
                                                                        RepFileProp[4] = "Aguarda armazenamento.";
                                                                        ImgFileProp[4] = "Aguarda armazenamento.";
                                                                        if (ImgFileProp[3] == "default" && ImgFileProp[3] != RepFileProp[3])
                                                                        { ImgFileProp[3] = RepFileProp[3]; }
                                                                        File.WriteAllLines(RepFilePath.Replace(".pdf", ".txt"), RepFileProp.ToArray());
                                                                        Thread.Sleep(100);
                                                                        File.WriteAllLines(ImgFilePath.Replace(".pdf", ".txt"), ImgFileProp.ToArray());
                                                                        Thread.Sleep(100);
                                                                        File.WriteAllLines(completefilename.Replace(".pdf", ".txt"), CompFileProp.ToArray());
                                                                        Thread.Sleep(100);
                                                                    }
                                                                    catch { erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao escrever os ficheiros identificadores de " + completefilename)); }
                                                                }
                                                                else
                                                                { erros.Add(Myhelp.Formatmsg("Não foi possível criar com sucesso o ficheiro " + completefilename)); }
                                                            }
                                                        }

                                                    }

                                                }
                                                else
                                                {
                                                    if (RepFileProp.Count > 0 && CompFileProp.Count == 0)
                                                    {
                                                        string RepFilePath = pathtoanalyze + @"\" + FileProp[6] + ".pdf";
                                                        if (Myhelp.FileInUse(RepFilePath) == false)
                                                        {
                                                            string targetfile = a.DirectoryName + @"\" + FileProp[5] + "C";
                                                            int x = 0;
                                                            if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                            {
                                                                while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                                                { x++; }
                                                            }
                                                            string completefilename = targetfile + x.ToString() + ".pdf";
                                                            File.Copy(RepFilePath, completefilename);
                                                            CompFileProp.AddRange(FileProp);
                                                            CompFileProp[2] = "C";
                                                            CompFileProp[4] = "Aguarda impressao.";
                                                            CompFileProp[8] = FileProp[5] + "C" + x.ToString();
                                                            FileProp[4] = "Aguarda armazenamento.";
                                                            try
                                                            {
                                                                File.WriteAllLines(infoFile, FileProp.ToArray());
                                                                Thread.Sleep(100);
                                                                File.WriteAllLines(completefilename.Replace(".pdf", ".txt"), CompFileProp.ToArray());
                                                                Thread.Sleep(100);
                                                            }
                                                            catch { }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                            if (FileProp[4] == "Impresso.")
                                            #region Impresso
                                            {
                                                if (FileProp[2] == "R")
                                                {
                                                    if (internalTempjoin == true)
                                                    {
                                                        FileProp[4] = "Aguarda Imagem.";
                                                    }
                                                    else
                                                    { FileProp[4] = "Aguarda armazenamento."; }
                                                }
                                                if (FileProp[2] == "I")
                                                {
                                                    if (internalTempjoin == true)
                                                    {
                                                        FileProp[4] = "Aguarda Relatorio.";
                                                    }
                                                    else
                                                    {
                                                        FileProp[4] = "Aguarda armazenamento.";
                                                    }
                                                }
                                                if (FileProp[2] == "C")
                                                {
                                                    FileProp[4] = "Aguarda armazenamento.";
                                                }
                                                try
                                                { File.WriteAllLines(infoFile, FileProp); }
                                                catch { }
                                            }
                                            #endregion
                                            if (FileProp[4] == "Aguarda impressao.")
                                            #region Aguarda Impressão
                                            {
                                                bool toprint = false;
                                                string printer = "";
                                                if (FileProp[2] == "R")
                                                {
                                                    toprint = internalTempRPrint;
                                                    printer = internalTempRPrinter;
                                                }
                                                if (FileProp[2] == "I")
                                                {
                                                    toprint = internalTempIPrint;
                                                    printer = internalTempIPrinter;
                                                }
                                                if (FileProp[2] == "C")
                                                {
                                                    toprint = internalTempCPrint;
                                                    printer = internalTempCPrinter;
                                                }
                                                if (toprint == true && printer != "")
                                                {
                                                    if (File.Exists(internalSumatraPdfpath))
                                                    {
                                                        if (Myhelp.printerisvalid(printer))
                                                        {
                                                            if (printer == internalBookletPrinter)//Bookletprint
                                                            {
                                                                bool procrunning = Myhelp.isProcessRunning(internalBookletProcess);
                                                                if (procrunning == false)
                                                                {
                                                                    if (Directory.Exists(internalBookletAudit))
                                                                    {
                                                                        string AuditIdentifier = "Print-in." + DateTime.Now.ToString("yyyyMMdd") + ".xml";
                                                                        if (File.Exists(internalBookletAudit + @"\" + AuditIdentifier))
                                                                        {
                                                                            if (debug == true)
                                                                            {
                                                                                erros.Add(Myhelp.Formatmsg("Última impressão a :" + Myhelp.LastBookletPrintTime(internalBookletAudit).ToString() + " Com SecsTowait - " + Myhelp.LastBookletPrintTime(internalBookletAudit).AddSeconds(internalBookletSecsToWait).ToString() + " Now: " + DateTime.Now.ToString()));
                                                                            }
                                                                            if (Myhelp.LastBookletPrintTime(internalBookletAudit).AddSeconds(internalBookletSecsToWait) < DateTime.Now)
                                                                            {
                                                                                #region Impressão para Booklet
                                                                                Process p = new System.Diagnostics.Process();
                                                                                ProcessStartInfo startInfo = new ProcessStartInfo();
                                                                                p.StartInfo.FileName = internalSumatraPdfpath;
                                                                                if (internaldebug == true)
                                                                                { p.StartInfo.Arguments = "-exit-when-done -print-to \"" + printer + "\" " + a.FullName.Replace(".txt", ".pdf"); }
                                                                                else
                                                                                {
                                                                                    p.StartInfo.Arguments = "-silent -exit-when-done -print-to \"" + printer + "\" " + a.FullName.Replace(".txt", ".pdf");
                                                                                }
                                                                                p.Start();
                                                                                p.WaitForExit();
                                                                                erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name.Replace(".txt", ".pdf") + " foi enviado para Booklet."));
                                                                                Thread.Sleep(500);
                                                                                //MessageBox.Show(BookletAudit + " " + a.Name + DateTime.Now.ToString() + " " + internalBookletSecsToWait.ToString());
                                                                                bool FoundPrint = Myhelp.FoundPrintJob(BookletAudit, a.Name.Replace(".txt", ".pdf"), DateTime.Now, internalBookletSecsToWait);
                                                                                if (FoundPrint == true)
                                                                                {
                                                                                    erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name.Replace(".txt", ".pdf") + " foi encontrado em " + AuditIdentifier));
                                                                                    if (FileProp[2] == "R")
                                                                                    {
                                                                                        RepFileProp[4] = "Impresso.";
                                                                                        try
                                                                                        { File.WriteAllLines(a.FullName, RepFileProp); }
                                                                                        catch { }
                                                                                    }
                                                                                    if (FileProp[2] == "I")
                                                                                    {
                                                                                        ImgFileProp[4] = "Impresso.";
                                                                                        try
                                                                                        { File.WriteAllLines(a.FullName, ImgFileProp); }
                                                                                        catch { }
                                                                                    }
                                                                                    if (FileProp[2] == "C")
                                                                                    {
                                                                                        CompFileProp[4] = "Impresso.";
                                                                                        try
                                                                                        { File.WriteAllLines(a.FullName, CompFileProp); }
                                                                                        catch { }
                                                                                    }
                                                                                }
                                                                                else
                                                                                {
                                                                                    if (FileProp[2] == "R")
                                                                                    {
                                                                                        RepFileProp[4] = "Impresso.";
                                                                                        try
                                                                                        { File.WriteAllLines(a.FullName, RepFileProp); }
                                                                                        catch { }
                                                                                    }
                                                                                    if (FileProp[2] == "I")
                                                                                    {
                                                                                        ImgFileProp[4] = "Impresso.";
                                                                                        try
                                                                                        { File.WriteAllLines(a.FullName, ImgFileProp); }
                                                                                        catch { }
                                                                                    }
                                                                                    if (FileProp[2] == "C")
                                                                                    {
                                                                                        CompFileProp[4] = "Impresso.";
                                                                                        try
                                                                                        { File.WriteAllLines(a.FullName, CompFileProp); }
                                                                                        catch { }
                                                                                    } erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name.Replace(".txt", ".pdf") + " não foi encontrado em " + AuditIdentifier + " após impressão para Booklet."));
                                                                                }
                                                                                #endregion
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            #region Impressão para Booklet
                                                                            Process p = new System.Diagnostics.Process();
                                                                            ProcessStartInfo startInfo = new ProcessStartInfo();
                                                                            p.StartInfo.FileName = internalSumatraPdfpath;
                                                                            if (internaldebug == true)
                                                                            { p.StartInfo.Arguments = "-exit-when-done -print-to \"" + printer + "\" " + a.FullName.Replace(".txt", ".pdf"); }
                                                                            else
                                                                            {
                                                                                p.StartInfo.Arguments = "-silent -exit-when-done -print-to \"" + printer + "\" " + a.FullName.Replace(".txt", ".pdf");
                                                                            }
                                                                            p.Start();
                                                                            p.WaitForExit();
                                                                            erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name.Replace(".txt", ".pdf") + " foi enviado para Booklet."));
                                                                            Thread.Sleep(500);
                                                                            //MessageBox.Show(BookletAudit + " " + a.Name + DateTime.Now.ToString() + " " + internalBookletSecsToWait.ToString());
                                                                            bool FoundPrint = Myhelp.FoundPrintJob(BookletAudit, a.Name.Replace(".txt", ".pdf"), DateTime.Now, internalBookletSecsToWait);
                                                                            if (FoundPrint == true)
                                                                            {
                                                                                erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name.Replace(".txt", ".pdf") + " foi encontrado em " + AuditIdentifier));
                                                                                if (FileProp[2] == "R")
                                                                                {
                                                                                    RepFileProp[4] = "Impresso.";
                                                                                    try
                                                                                    { File.WriteAllLines(a.FullName, RepFileProp); }
                                                                                    catch { }
                                                                                }
                                                                                if (FileProp[2] == "I")
                                                                                {
                                                                                    ImgFileProp[4] = "Impresso.";
                                                                                    try
                                                                                    { File.WriteAllLines(a.FullName, ImgFileProp); }
                                                                                    catch { }
                                                                                }
                                                                                if (FileProp[2] == "C")
                                                                                {
                                                                                    CompFileProp[4] = "Impresso.";
                                                                                    try
                                                                                    { File.WriteAllLines(a.FullName, CompFileProp); }
                                                                                    catch { }
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                if (FileProp[2] == "R")
                                                                                {
                                                                                    RepFileProp[4] = "Impresso.";
                                                                                    try
                                                                                    { File.WriteAllLines(a.FullName, RepFileProp); }
                                                                                    catch { }
                                                                                }
                                                                                if (FileProp[2] == "I")
                                                                                {
                                                                                    ImgFileProp[4] = "Impresso.";
                                                                                    try
                                                                                    { File.WriteAllLines(a.FullName, ImgFileProp); }
                                                                                    catch { }
                                                                                }
                                                                                if (FileProp[2] == "C")
                                                                                {
                                                                                    CompFileProp[4] = "Impresso.";
                                                                                    try
                                                                                    { File.WriteAllLines(a.FullName, CompFileProp); }
                                                                                    catch { }
                                                                                } erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name.Replace(".txt", ".pdf") + " não foi encontrado em " + AuditIdentifier + " após impressão para Booklet."));
                                                                            }
                                                                            #endregion

                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                #region Impressão normal

                                                                Process p = new System.Diagnostics.Process();
                                                                ProcessStartInfo startInfo = new ProcessStartInfo();
                                                                p.StartInfo.FileName = internalSumatraPdfpath;
                                                                p.StartInfo.Arguments = "-silent -exit-when-done -print-to \"" + printer + "\" " + a.FullName.Replace(".txt", ".pdf"); ;
                                                                p.Start();
                                                                p.WaitForExit();
                                                                erros.Add(Myhelp.Formatmsg("O ficheiro " + a.Name.Replace(".txt", ".pdf") + "  foi enviado para " + printer));
                                                                Thread.Sleep(500);
                                                                if (FileProp[2] == "R")
                                                                {
                                                                    RepFileProp[4] = "Impresso.";
                                                                    try
                                                                    { File.WriteAllLines(a.FullName, RepFileProp); }
                                                                    catch { }
                                                                }
                                                                if (FileProp[2] == "I")
                                                                {
                                                                    ImgFileProp[4] = "Impresso.";
                                                                    try
                                                                    { File.WriteAllLines(a.FullName, ImgFileProp); }
                                                                    catch { }
                                                                }
                                                                if (FileProp[2] == "C")
                                                                {
                                                                    CompFileProp[4] = "Impresso.";
                                                                    try
                                                                    { File.WriteAllLines(a.FullName, CompFileProp); }
                                                                    catch { }
                                                                }
                                                                #endregion

                                                            }//print
                                                        }
                                                        else { erros.Add(Myhelp.Formatmsg(" A impressora " + printer + " não é valida ou não está disponivel.")); }
                                                    }
                                                    else { erros.Add(Myhelp.Formatmsg(" O ficheiro " + internalSumatraPdfpath + " não existe.")); }
                                                }
                                                else
                                                {
                                                    if (toprint == false)
                                                    {
                                                        if (FileProp[2] == "R")
                                                        {
                                                            RepFileProp[4] = "Impresso.";
                                                            try
                                                            { File.WriteAllLines(a.FullName, RepFileProp); }
                                                            catch { }
                                                        }
                                                        if (FileProp[2] == "I")
                                                        {
                                                            ImgFileProp[4] = "Impresso.";
                                                            try
                                                            { File.WriteAllLines(a.FullName, ImgFileProp); }
                                                            catch { }
                                                        }
                                                        if (FileProp[2] == "C")
                                                        {
                                                            CompFileProp[4] = "Impresso.";
                                                            try
                                                            { File.WriteAllLines(a.FullName, CompFileProp); }
                                                            catch { }
                                                        }


                                                    }
                                                }
                                            }
                                            #endregion

                                        }
                                    }
                                    else { erros.Add(Myhelp.Formatmsg(" O ficheiro " + infoFile + " não é um ficheiro identificador válido.")); }
                                    ProcessedFiles.Add(infoFile.Replace(".txt", ""));
                                }
                                else
                                {
                                    if (File.Exists(pdfFile) == false)
                                    {
                                        try
                                        {
                                            File.Delete(infoFile);
                                            erros.Add(Myhelp.Formatmsg(" O ficheiro " + infoFile + " foi apagado."));
                                        }
                                        catch { erros.Add(Myhelp.Formatmsg(" Não foi possível apagar o ficheiro " + infoFile + " .")); }
                                    }
                                }
                            }
                        }
                    }
                    else { erros.Add(Myhelp.Formatmsg(" A directoria " + pathtoanalyze + " não existe ou não está disponível.")); }
                }
            } */
                    #endregion

                    if (erros.Count > 0)
                    {
                        internallogger.AddEx(erros);
                    }
                   
                }
            }
        }
        private void BCKPDFOutput_DoWork(object sender, DoWorkEventArgs e)
        {
            Helper Myhelp = new Helper();
            #region Variaveis
            int error = 0;
            
            List<string> erros = new List<string>();
            List<string> FullTbl = new List<string>();
            List<string> PDFContent = new List<string>();
            List<string> ReportContent = new List<string>();
            List<string> ImgContent = new List<string>();
            object[] parameters = e.Argument as object[];
            ListLogger internallogger = (ListLogger)parameters[0];
            string internalpathPDF = "";
            string internalpathECG = "";
            string internalpathECGTbl = "";
            string internalpathECO = "";
            string internalpathECOTbl = "";
            string internalpathHOLTER = "";
            string internalpathHOLTERTbl = "";
            string internalpathMAPA = "";
            string internalpathMAPATbl = "";
            string internalpathPE = "";
            string internalpathPETbl = "";
            string internalpathUnconfirmed = "";
            try
            {
                internalpathPDF = (string)parameters[1];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathPDF em BCKPDFOutput.")); }
            try
            {
                internalpathECG = (string)parameters[2];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathECG em BCKPDFOutput.")); }
            try
            {
                internalpathECGTbl = (string)parameters[3];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathECGTbl em BCKPDFOutput.")); }
            try
            {
                internalpathECO = (string)parameters[4];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathECO em BCKPDFOutput.")); }
            try
            {
                internalpathECOTbl = (string)parameters[5];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathECOTbl em BCKPDFOutput.")); }
            try
            {
                internalpathHOLTER = (string)parameters[6];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathHOLTER em BCKPDFOutput.")); }
            try
            {
                internalpathHOLTERTbl = (string)parameters[7];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathHOLTERTbl em BCKPDFOutput.")); }
            try
            {
                internalpathMAPA = (string)parameters[8];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathMAPA em BCKPDFOutput.")); }
            try
            {
                internalpathMAPATbl = (string)parameters[9];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathMAPATbl em BCKPDFOutput.")); }
            try
            {
                internalpathPE = (string)parameters[10];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathPE em BCKPDFOutput.")); }
            try
            {
                internalpathPETbl = (string)parameters[11];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathPETbl em BCKPDFOutput.")); }
            try
            {
                internalpathUnconfirmed = (string)parameters[12];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathUnconfirmed em BCKPDFOutput.")); } 
            #endregion
            List<string> MyTables = new List<string>();
            if (File.Exists(internalpathECGTbl))
            { MyTables.Add(internalpathECGTbl); }
            if (File.Exists(internalpathECOTbl))
            { MyTables.Add(internalpathECOTbl); }
            if (File.Exists(internalpathHOLTERTbl))
            { MyTables.Add(internalpathHOLTERTbl); }
            if (File.Exists(internalpathMAPATbl))
            { MyTables.Add(internalpathMAPATbl); }
            if (File.Exists(internalpathPETbl))
            { MyTables.Add(internalpathPETbl); }
            if (error == 0)
            {
                try
                {
                    List<string> dir = Directory.GetFiles(internalpathPDF,"*.pdf").ToList();
                    if (dir.Count > 0)
                    {
                        foreach (string s in dir)
                        {
                            if (File.Exists(s))
                            {
                                erros.Add(Myhelp.Formatmsg("A analisar: " + s));
                                if (Myhelp.FileInUse(s) == false)
                                {

                                    if (MyTables.Count > 0)
                                    {

                                        string Temptbl = "";
                                        string TargetPath = "";
                                        PDFContent = Myhelp.ReadPdfFile(s);
                                        int ThisExamTyp = Myhelp.ExamType(PDFContent, MyTables);
                                        List<string> FileProp = new List<string>();
                                        FileProp.Add("default");//Nome
                                        FileProp.Add("default");//ID
                                        FileProp.Add("default");//Tipo de exame (R ou I ou C)
                                        FileProp.Add("default");//Data do exame
                                        FileProp.Add("default");//Estado deste ficheiro
                                        FileProp.Add("default");//Nome do Exame
                                        FileProp.Add("default");//Nome deste ficheiro
                                        FileProp.Add("default");//Nome do ficheiro partner
                                        FileProp.Add("default");//Nome do ficheiro completo    
                                        FileInfo a = new FileInfo(s);
                                        if (ThisExamTyp == 1 || ThisExamTyp == 2)
                                        {
                                            Temptbl = internalpathECGTbl;
                                            TargetPath = internalpathECG;
                                            FileProp[5] = "ECG";
                                            if (ThisExamTyp == 1)
                                            { FileProp[2] = "R"; erros.Add(Myhelp.Formatmsg(s + " identificado como relatório de ECG.")); }
                                            else { FileProp[2] = "I"; erros.Add(Myhelp.Formatmsg(s + " identificado como suportes de ECG.")); }

                                        }
                                        if (ThisExamTyp == 3 || ThisExamTyp == 4)
                                        {
                                            Temptbl = internalpathECOTbl;
                                            TargetPath = internalpathECO;
                                            FileProp[5] = "ECO";
                                            if (ThisExamTyp == 3)
                                            { FileProp[2] = "R"; erros.Add(Myhelp.Formatmsg(s + " identificado como relatório de ECO.")); }
                                            else { FileProp[2] = "I"; erros.Add(Myhelp.Formatmsg(s + " identificado como suportes de ECO.")); }

                                        }
                                        if (ThisExamTyp == 5 || ThisExamTyp == 6)
                                        {
                                            Temptbl = internalpathHOLTERTbl;
                                            TargetPath = internalpathHOLTER;
                                            FileProp[5] = "HOLTER";
                                            if (ThisExamTyp == 5)
                                            { FileProp[2] = "R"; erros.Add(Myhelp.Formatmsg(s + " identificado como relatório de HOLTER.")); }
                                            else { FileProp[2] = "I"; erros.Add(Myhelp.Formatmsg(s + " identificado como suportes de HOLTER.")); }

                                        }
                                        if (ThisExamTyp == 7 || ThisExamTyp == 8)
                                        {
                                            Temptbl = internalpathMAPATbl;
                                            TargetPath = internalpathMAPA;
                                            FileProp[5] = "MAPA";
                                            if (ThisExamTyp == 7)
                                            { FileProp[2] = "R"; erros.Add(Myhelp.Formatmsg(s + " identificado como relatório de MAPA.")); }
                                            else { FileProp[2] = "I"; erros.Add(Myhelp.Formatmsg(s + " identificado como suportes de MAPA.")); }

                                        }
                                        if (ThisExamTyp == 9 || ThisExamTyp == 10)
                                        {
                                            Temptbl = internalpathPETbl;
                                            TargetPath = internalpathPE;
                                            FileProp[5] = "PE";
                                            if (ThisExamTyp == 9)
                                            { FileProp[2] = "R"; erros.Add(Myhelp.Formatmsg(s + " identificado como relatório de PE.")); }
                                            else { FileProp[2] = "I"; erros.Add(Myhelp.Formatmsg(s + " identificado como suportes de PE.")); }

                                        }
                                        if (ThisExamTyp == 11 || ThisExamTyp == 12 || ThisExamTyp == 13 || ThisExamTyp == 14)
                                        {

                                            TargetPath = internalpathUnconfirmed;
                                            erros.Add(Myhelp.Formatmsg("Não foi possivel verificar o tipo de exame do ficheiro " + s));
                                            FileProp[5] = "UC";
                                            FileProp[2] = "UC";

                                        }
                                        if (ThisExamTyp < 11)
                                        {
                                            FileProp[0] = Myhelp.GetNormalizedName(Myhelp.GetName(ThisExamTyp, PDFContent, Temptbl));
                                            FileProp[1] = Myhelp.GetNormalizedID(Myhelp.GetId(ThisExamTyp, PDFContent, Temptbl));
                                            FileProp[3] = Myhelp.GetDate(ThisExamTyp, PDFContent, Temptbl);
                                            FileProp[4] = "Aguarda impressao.";
                                        }
                                        else
                                        {
                                            FileProp[0] = "Desconhecido";
                                            FileProp[1] = "Desconhecido";
                                            FileProp[3] = "Desconhecido";
                                            FileProp[4] = "Não definido.";
                                        }
                                        string targetfile = TargetPath + @"\" + FileProp[5] + FileProp[2];
                                        int x = 0;
                                        if (File.Exists(targetfile + x.ToString() + ".pdf"))
                                        {
                                            while (File.Exists(targetfile + x.ToString() + ".pdf"))
                                            { x++; }
                                        }
                                        if (FileProp[2] == "R")
                                        {
                                            FileProp[6] = FileProp[5] + FileProp[2] + x.ToString();
                                        }
                                        if (FileProp[2] == "I")
                                        {
                                            FileProp[7] = FileProp[5] + FileProp[2] + x.ToString();
                                        }
                                        if (FileProp[2] == "UC")
                                        {
                                            FileProp[6] = FileProp[5] + FileProp[2] + x.ToString();
                                        }
                                        bool moveeror = false;
                                        try
                                        {
                                            a.MoveTo(targetfile + x.ToString() + ".pdf");
                                        }
                                        catch { erros.Add(Myhelp.Formatmsg("Não foi possível mover o ficheiro: " + a.FullName)); moveeror = true; }
                                        if (moveeror == false)
                                        {
                                            try
                                            {
                                                if (Directory.Exists(TargetPath))
                                                {
                                                    try
                                                    {
                                                        File.WriteAllLines(targetfile + x.ToString() + ".txt", FileProp.ToArray());
                                                    }
                                                    catch { erros.Add(Myhelp.Formatmsg("Não foi possível criar o ficheiro: " + targetfile + x.ToString() + ".txt")); }
                                                    erros.Add(Myhelp.Formatmsg(s + " movido para " + TargetPath + " com o nome:" + targetfile + x.ToString() + ".pdf"));
                                                }
                                            }
                                            catch
                                            {
                                                erros.Add(Myhelp.Formatmsg("Não foi possível mover " + s + " para " + TargetPath));
                                            }
                                        }
                                    }
                                    else
                                    { erros.Add(Myhelp.Formatmsg("Não foi possível analisar " + s + " porque não há tabelas de comparação.")); }
                                }
                                else { erros.Add(Myhelp.Formatmsg("Não foi possível analisar " + s + " porque o ficheiro está a ser utilizado.")); }
                            }
                            else { erros.Add(Myhelp.Formatmsg("Ficheiro " + s + " não existe.")); }
                        
                        }
                    }
                }
                catch { erros.Add(Myhelp.Formatmsg("Ocorreu um erro geral em BCKPDFOutput.")); }
            }
            if (erros.Count > 0)
            {
                internallogger.AddEx(erros);
            }
        }
        private void BCKUpdateLst_DoWork(object sender, DoWorkEventArgs e)
        {
            #region Variaveis
            Helper Myhelp = new Helper();
            List<string> erros = new List<string>();
            int error = 0;
            object[] parameters = e.Argument as object[];
            ListLogger internallogger = (ListLogger)parameters[0];
            string internalpathECG = "";
            string internalpathECO = "";
            string internalpathHOLTER = "";
            string internalpathMAPA = "";
            string internalpathPE = "";
            string internalpathUnconfirmed = "";
            int internalstate = 2;
            bool internalfirstload = new bool();
            List<TreeNode> internaloldECG = new List<TreeNode>();
            List<TreeNode> internaloldECOlst = new List<TreeNode>();
            List<TreeNode> internaloldHOLTERlst = new List<TreeNode>();
            List<TreeNode> internaloldMAPAlst = new List<TreeNode>();
            List<TreeNode> internaloldPElst = new List<TreeNode>();
            List<TreeNode> internaloldUnconfirmedlst = new List<TreeNode>();
            try { internalstate = (int)parameters[14]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalstate em BCKUpdateLst.")); }
            try { internalfirstload = (bool)parameters[13]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalfirstload em BCKUpdateLst.")); }
            try { internaloldECG = (List<TreeNode>)parameters[7]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaloldECGlst em BCKUpdateLst.")); }
            try { internaloldECOlst = (List<TreeNode>)parameters[8]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaloldECOlst em BCKUpdateLst.")); }

            try { internaloldHOLTERlst = (List<TreeNode>)parameters[9]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaloldHOLTERlst em BCKUpdateLst.")); }
            try { internaloldMAPAlst = (List<TreeNode>)parameters[10]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaloldMAPAlst em BCKUpdateLst.")); }
            try { internaloldPElst = (List<TreeNode>)parameters[11]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaloldPElst em BCKUpdateLst.")); }
            try { internaloldUnconfirmedlst = (List<TreeNode>)parameters[12]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaloldUnconfirmedlst em BCKUpdateLst.")); }
            
            try
            {
                internalpathECG = (string)parameters[1];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathECG em BCKUpdateLst.")); }
            try
            {
                internalpathECO = (string)parameters[2];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathECO em BCKUpdateLst.")); }
            try
            {
                internalpathHOLTER = (string)parameters[3];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathHOLTER em BCKUpdateLst.")); }
            try
            {
                internalpathMAPA = (string)parameters[4];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathMAPA em BCKUpdateLst.")); }
            try
            {
                internalpathPE = (string)parameters[5];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathPE em BCKUpdateLst.")); }
            try
            {
                internalpathUnconfirmed = (string)parameters[6];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathUnconfirmed em BCKUpdateLst.")); }
            if (error == 0)
            {

                List<string> analyzingpatf = new List<string>();
                analyzingpatf.Add(internalpathECG);
                analyzingpatf.Add(internalpathECO);
                analyzingpatf.Add(internalpathHOLTER);
                analyzingpatf.Add(internalpathMAPA);
                analyzingpatf.Add(internalpathPE);
                analyzingpatf.Add(internalpathUnconfirmed);

                for (int i = 0; i <= analyzingpatf.Count - 1; i++)
                {
                    string pathtoanalyze = analyzingpatf[i];
                    if (Directory.Exists(pathtoanalyze))
                    {
                        List<Prijob> Jobs = new List<Prijob>();
                        List<string> PathFiles = Directory.GetFiles(pathtoanalyze, "*.txt", SearchOption.TopDirectoryOnly).ToList();
                        if (PathFiles.Count > 0)
                        {
                            foreach (string txtfile in PathFiles)
                            #region Criar Jobs
                            {

                                string associatePdf = txtfile.Replace(".txt", ".pdf");
                                if (File.Exists(associatePdf))
                                {
                                    
                                       
                                            
                                                List<string> content = new List<string>();
                                                try
                                                {
                                                    content = File.ReadAllLines(txtfile).ToList();
                                                }
                                                catch { erros.Add(Myhelp.Formatmsg("Não foi possível ler o ficheiro " + txtfile)); }//mensagem de erro na leitura do ficheiro contente
                                                if (content.Count >= 8)
                                                {
                                                    if (Jobs.Count == 0)
                                                    {
                                                        Prijob newjob = new Prijob(content[0], content[1]);
                                                        if (content[2] == "R")
                                                        { newjob.AddReport(content); }
                                                        if (content[2] == "I")
                                                        { newjob.AddImage(content); }
                                                        if (content[2] == "C")
                                                        { newjob.AddComplete(content); }
                                                        Jobs.Add(newjob);
                                                    }
                                                    else
                                                    {
                                                        if (Jobs.Any(job => job.GId == content[1]))
                                                        {
                                                            foreach (Prijob Job in Jobs)
                                                            {
                                                                if (Job.GId == content[1])
                                                                {
                                                                    if (content[2] == "R")
                                                                    { Job.AddReport(content); }
                                                                    if (content[2] == "I")
                                                                    { Job.AddImage(content); }
                                                                    if (content[2] == "C")
                                                                    { Job.AddComplete(content); }
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Prijob newjob = new Prijob(content[0], content[1]);
                                                            if (content[2] == "R")
                                                            { newjob.AddReport(content); }
                                                            if (content[2] == "I")
                                                            { newjob.AddImage(content); }
                                                            if (content[2] == "C")
                                                            { newjob.AddComplete(content); }
                                                            Jobs.Add(newjob);
                                                        }
                                                    }
                                                }
                                                else { erros.Add(Myhelp.Formatmsg(txtfile + " é um ficheiro inválido.")); }
                                            
                                        
                                    
                                }
                                
                            }
                            #endregion

                        }
                        List<TreeNode> MasterList = new List<TreeNode>();
                        
                        
                        foreach (Prijob Job in Jobs)
                        {
                            TreeNode Master = new TreeNode();
                            Master.Name = Job.Gname + Job.GId;
                            Master.Text = Job.Gname + " (" + Job.GId + ")";

                            TreeNode MasterRep = new TreeNode();
                            MasterRep.Name = "Reports";
                            MasterRep.Text = "Relatórios";
                            Master.Nodes.Add(MasterRep);
                            TreeNode MasterImg = new TreeNode();
                            MasterImg.Name = "Images";
                            MasterImg.Text = "Imagens";
                            Master.Nodes.Add(MasterImg);
                            TreeNode MasterCpt = new TreeNode();
                            MasterCpt.Name = "Complete";
                            MasterCpt.Text = "Completos";
                            Master.Nodes.Add(MasterCpt);
                            List<List<string>> rep = Job.GReports;
                            List<List<string>> img = Job.GImages;
                            List<List<string>> cpt = Job.GComplete;
                            foreach (List<string> rp in rep)
                            {
                                if (rp.Count >= 8)
                                {
                                    TreeNode ChildNode = new TreeNode();
                                    ChildNode.Name = rp[6];
                                    ChildNode.Text = rp[6];
                                    TreeNode SubChildId = new TreeNode();
                                    TreeNode SubchildState = new TreeNode();
                                    SubChildId.Name = "ID";
                                    SubChildId.Text = rp[1];
                                    ChildNode.Nodes.Add(SubChildId);
                                    SubchildState.Name = "State";
                                    SubchildState.Text = rp[4];
                                    if (rp[4] == "Armazenado." || rp[4] == "Aguarda armazenamento.")
                                    { ChildNode.BackColor = Color.LimeGreen;
                                    if (Master.BackColor == Color.LightYellow || Master.BackColor == Color.LightBlue || Master.BackColor == Color.IndianRed)
                                    {  }
                                    else
                                    { Master.BackColor = Color.LimeGreen; }
                                    }
                                    if (rp[4] == "Aguarda Imagem." || rp[4] == "Aguarda impressao." || rp[4] == "Impresso.")
                                    {
                                        ChildNode.BackColor = Color.LightYellow;
                                        
                                         Master.BackColor = Color.LightYellow; 
                                    }
                                    if (rp[4] == "reset")
                                    { ChildNode.BackColor = Color.LightBlue;
                                   
                                     Master.BackColor = Color.LightBlue; 
                                    }
                                    if (rp[4] == "Apagar")
                                    { ChildNode.BackColor = Color.IndianRed;
                                    
                                     Master.BackColor = Color.IndianRed; 
                                    }
                                    ChildNode.Nodes.Add(SubchildState);
                                    MasterRep.Nodes.Add(ChildNode);
                                    
                                }
                            }
                            foreach (List<string> im in img)
                            {
                                if (im.Count >= 8)
                                {
                                    TreeNode ChildNode = new TreeNode();
                                    ChildNode.Name = im[7];
                                    ChildNode.Text = im[7];
                                    TreeNode SubChildId = new TreeNode();
                                    TreeNode SubchildState = new TreeNode();
                                    SubChildId.Name = "ID";
                                    SubChildId.Text = im[1];
                                    ChildNode.Nodes.Add(SubChildId);
                                    SubchildState.Name = "State";
                                    SubchildState.Text = im[4];
                                    if (im[4] == "Armazenado." || im[4] == "Aguarda armazenamento.")
                                    { ChildNode.BackColor = Color.LimeGreen;
                                    if (Master.BackColor == Color.LightYellow || Master.BackColor == Color.LightBlue || Master.BackColor == Color.IndianRed)
                                    {  }
                                    else

                                    { Master.BackColor = Color.LimeGreen; }
                                    }
                                    if (im[4] == "Aguarda Relatorio." || im[4] == "Aguarda impressao." || im[4] == "Impresso.")
                                    { ChildNode.BackColor = Color.LightYellow;
                                    
                                    Master.BackColor = Color.LightYellow; 
                                    }
                                    if (im[4] == "reset")
                                    { ChildNode.BackColor = Color.LightBlue;
                                    
                                     Master.BackColor = Color.LightBlue; 
                                    }
                                    if (im[4] == "Apagar")
                                    { ChildNode.BackColor = Color.IndianRed;
                                    
                                     Master.BackColor = Color.IndianRed; 
                                    }
                                    ChildNode.Nodes.Add(SubchildState);
                                    MasterImg.Nodes.Add(ChildNode);
                                    
                                }
                            }
                            foreach (List<string> cp in cpt)
                            {
                                if (cp.Count >= 8)
                                {
                                    TreeNode ChildNode = new TreeNode();
                                    ChildNode.Name = cp[8];
                                    ChildNode.Text = cp[8];
                                    TreeNode SubChildId = new TreeNode();
                                    TreeNode SubchildState = new TreeNode();
                                    SubChildId.Name = "ID";
                                    SubChildId.Text = cp[1];
                                    ChildNode.Nodes.Add(SubChildId);
                                    SubchildState.Name = "State";
                                    SubchildState.Text = cp[4];
                                    if (cp[4] == "Aguarda impressao." || cp[4] == "Impresso.")
                                    {
                                        ChildNode.BackColor = Color.LightYellow;

                                        Master.BackColor = Color.LightYellow;
                                    }
                                    if (cp[4] == "Armazenado." || cp[4] == "Aguarda armazenamento.")
                                    { ChildNode.BackColor = Color.LimeGreen;
                                    if (Master.BackColor == Color.LightYellow || Master.BackColor == Color.LightBlue || Master.BackColor == Color.IndianRed)
                                    {  }
                                    else
                                    { Master.BackColor = Color.LimeGreen; }
                                    }
                                    else
                                    {
                                        if (cp[4] == "reset")
                                        { ChildNode.BackColor = Color.LightBlue;
                                        
                                         Master.BackColor = Color.LightBlue; 
                                        }
                                        else
                                        { ChildNode.BackColor = Color.LightYellow;
                                        
                                        Master.BackColor = Color.LightYellow; 
                                        }
                                    }
                                     if (cp[4] == "Apagar")
                                        { ChildNode.BackColor = Color.IndianRed;
                                        
                                         Master.BackColor = Color.IndianRed; 
                                     }
                                    
                                    ChildNode.Nodes.Add(SubchildState);
                                    MasterCpt.Nodes.Add(ChildNode);
                                    
                                    
                                }
                            }

                            MasterList.Add(Master);

                        }
                        List<string> Masterlines = new List<string>();
                        List<string> oldMasterLines = new List<string>();
                        
                        string masterlines = "";
                        foreach (TreeNode a in MasterList)
                        {
                            masterlines = "";
                            masterlines = masterlines + a.Text + a.BackColor;
                            foreach (TreeNode b in a.Nodes)
                            {
                                masterlines = masterlines + b.Text;
                                foreach (TreeNode c in b.Nodes)
                                {
                                    if (File.Exists(pathtoanalyze + @"\" + c.Text + ".txt"))
                                    { FileInfo tocheck = new FileInfo(pathtoanalyze + @"\" + c.Text + ".txt");
                                    if (tocheck.LastWriteTime.ToShortDateString() == DateTime.Now.ToShortDateString())
                                    {
                                        if (a.BackColor != Color.LimeGreen)
                                        {
                                            a.Expand(); b.Expand();
                                        }
                                    }
                                    }
                                    masterlines = masterlines + c.Text + c.BackColor.ToString();
                                foreach (TreeNode d in c.Nodes)
                                {
                                    masterlines = masterlines + d.Text;

                                }
                                }
                                
                            }
                            Masterlines.Add(masterlines);
                        }
                        if (File.Exists(pathtoanalyze + @"\" + "index.tbl"))
                        { try { oldMasterLines = File.ReadAllLines(pathtoanalyze + @"\" + "index.tbl").ToList(); } catch { } }
                        else { try {
                            if (internalstate == 0)
                            {
                                File.WriteAllLines(pathtoanalyze + @"\" + "index.tbl", Masterlines.ToArray());
                            }
                        } catch { } }
 
                        var listsequal = Masterlines.All(oldMasterLines.Contains) && Masterlines.Count == oldMasterLines.Count;
                        if (internalfirstload == true)
                        {
                            try
                            {
                                if (internalstate == 0)
                                {
                                    File.WriteAllLines(pathtoanalyze + @"\" + "index.tbl", Masterlines.ToArray());
                                }
                            }
                            catch { }
                            object[] parameters2 = new object[] { MasterList };
                            BCKUpdateLst.ReportProgress(i, parameters2);
                        }
                        else
                        {
                            if (listsequal == false)
                            {
                                try
                                {
                                    if (internalstate == 0)
                                    {
                                        File.WriteAllLines(pathtoanalyze + @"\" + "index.tbl", Masterlines.ToArray());
                                    }
                                }
                                catch { }
                                object[] parameters2 = new object[] { MasterList };
                                BCKUpdateLst.ReportProgress(i, parameters2);
                            }
                        }
                    }
                }
            }
            #endregion
            #region OldUpdateLst
            /*
            #region Variaveis
            Helper Myhelp = new Helper();
            List<string> erros = new List<string>();
            int error = 0;
            object[] parameters = e.Argument as object[];
            ListLogger internallogger = (ListLogger)parameters[0];
            string internalpathECG = "";
            string internalpathECO = "";
            string internalpathHOLTER = "";
            string internalpathMAPA = "";
            string internalpathPE = "";
            string internalpathUnconfirmed = "";
            List<string> internaloldECGlst = new List<string>();
            List<string> internaloldECOlst = new List<string>();
            List<string> internaloldHOLTERlst = new List<string>();
            List<string> internaloldMAPAlst = new List<string>();
            List<string> internaloldPElst = new List<string>();
            List<string> internaloldUnconfirmedlst = new List<string>();
            try { internaloldECGlst = (List<string>)parameters[7]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaloldECGlst em BCKUpdateLst.")); }
            try { internaloldECOlst = (List<string>)parameters[8]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaloldECOlst em BCKUpdateLst.")); }
            try { internaloldHOLTERlst = (List<string>)parameters[9]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaloldHOLTERlst em BCKUpdateLst.")); }
            try { internaloldMAPAlst = (List<string>)parameters[10]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaloldMAPAlst em BCKUpdateLst.")); }
            try { internaloldPElst = (List<string>)parameters[11]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaloldPElst em BCKUpdateLst.")); }
            try { internaloldUnconfirmedlst = (List<string>)parameters[12]; }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internaloldUnconfirmedlst em BCKUpdateLst.")); }
            try
            {
                internalpathECG = (string)parameters[1];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathECG em BCKUpdateLst.")); }
            try
            {
                internalpathECO = (string)parameters[2];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathECO em BCKUpdateLst.")); }
            try
            {
                internalpathHOLTER = (string)parameters[3];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathHOLTER em BCKUpdateLst.")); }
            try
            {
                internalpathMAPA = (string)parameters[4];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathMAPA em BCKUpdateLst.")); }
            try
            {
                internalpathPE = (string)parameters[5];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathPE em BCKUpdateLst.")); }
            try
            {
                internalpathUnconfirmed = (string)parameters[6];

            }
            catch { error++; erros.Add(Myhelp.Formatmsg("Ocorreu um erro ao atribuir valores a internalpathUnconfirmed em BCKUpdateLst.")); }

            #endregion
            if (error == 0)
            {

                List<string> analyzingpatf = new List<string>();              
                List<ListViewItem> newlist = new List<ListViewItem>();
                analyzingpatf.Add(internalpathECG);
                analyzingpatf.Add(internalpathECO);
                analyzingpatf.Add(internalpathHOLTER);
                analyzingpatf.Add(internalpathMAPA);
                analyzingpatf.Add(internalpathPE);
                analyzingpatf.Add(internalpathUnconfirmed);

                for (int i = 0; i <= analyzingpatf.Count - 1; i++)
                {
                    List<string> ProcessedFiles = new List<string>();
                    Thread.Sleep(500);
                    List<string> oldlist = new List<string>();
                    if (i == 0)
                    { oldlist = internaloldECGlst; }
                    if (i == 1)
                    { oldlist = internaloldECOlst; }
                    if (i == 2)
                    { oldlist = internaloldHOLTERlst; }
                    if (i == 3)
                    { oldlist = internaloldMAPAlst; }
                    if (i == 4)
                    { oldlist = internaloldPElst; }
                    if (i == 5)
                    { oldlist = internaloldUnconfirmedlst;  }
                    newlist.Clear();
                    string pathtoanalyze = analyzingpatf[i];
                    bool dontjoin = false;
                    
                    if (Directory.Exists(pathtoanalyze))
                    {
                        DirectoryInfo di = new DirectoryInfo(pathtoanalyze);
                        List<FileInfo> PathFiles = di.GetFiles("*.txt", SearchOption.TopDirectoryOnly).ToList();
                        List<FileInfo> orderedList = PathFiles.OrderByDescending(x => x.CreationTime).ToList();
                        foreach (FileInfo a in orderedList)
                        {
                            dontjoin = false;
                            if (File.Exists(a.FullName.Replace(".txt", ".pdf")) && ProcessedFiles.Contains(a.Name.Replace(".txt","")) == false)
                            {
                                List<string> FileCt = new List<string>();
                                FileCt = File.ReadAllLines(a.FullName).ToList();
                                if (FileCt.Count >= 8)
                                {
                                    ListViewItem newExamRow = new ListViewItem(FileCt[0]);
                                    newExamRow.SubItems.Add(FileCt[1]);
                                    //Relatórios e Imagens
                                    
                                    if (FileCt[6] != "default")
                                    { newExamRow.SubItems.Add("Sim"); ProcessedFiles.Add(FileCt[6]);}
                                    else
                                    { newExamRow.SubItems.Add("Não"); }
                                    if (FileCt[7] != "default" && FileCt[7] != "false")
                                    { newExamRow.SubItems.Add("Sim"); ProcessedFiles.Add(FileCt[7]); }
                                    else
                                    {
                                        if (FileCt[7] == "false")
                                        { newExamRow.SubItems.Add("Desnecessário"); }
                                        else
                                        { newExamRow.SubItems.Add("Não"); }
                                    }
                                    
                                        newExamRow.SubItems.Add(FileCt[3]);
                                    
                                    //Estados
                                    List<string> FileCtRep = new List<string>();
                                    List<string> FileCtImg = new List<string>();
                                    List<string> FileCtCpt = new List<string>();
                                    if (FileCt[6] != "default" && File.Exists(a.Directory + @"\" + FileCt[6] + ".txt"))
                                    { FileCtRep = File.ReadAllLines(a.Directory + @"\" + FileCt[6] + ".txt").ToList(); }
                                    if (FileCt[7] != "default" && File.Exists(a.Directory + @"\" + FileCt[7] + ".txt"))
                                    { FileCtImg = File.ReadAllLines(a.Directory + @"\" + FileCt[7] + ".txt").ToList(); }
                                    if (FileCt[8] != "default" && File.Exists(a.Directory + @"\" + FileCt[8] + ".txt"))
                                    { FileCtCpt = File.ReadAllLines(a.Directory + @"\" + FileCt[8] + ".txt").ToList(); }
                                    if (FileCt[2] == "UC")
                                    { newExamRow.SubItems.Add("Desconhecido"); }
                                    else
                                    {
                                        if (FileCtRep.Count >= 8 && FileCtImg.Count >= 8 && FileCtCpt.Count >= 8)
                                        {
                                            if (FileCtRep[4] == FileCtImg[4] && FileCtImg[4] == FileCtCpt[4])
                                            { newExamRow.SubItems.Add(FileCtRep[4]); }
                                            else
                                            { newExamRow.SubItems.Add("Rel: " + FileCtRep[4] + " " + "Img: " + FileCtImg[4] + " " + "Completo: " + FileCtCpt[4]); }
                                        }
                                        else
                                        {
                                            string RelStat = "";
                                            string ImgStat = "";
                                            string CptStat = "";
                                            if (FileCtRep.Count >= 8 && FileCtImg.Count == 0)
                                            { newExamRow.SubItems.Add(FileCtRep[4]); }
                                            else
                                            {
                                                if (FileCtRep.Count == 0 && FileCtImg.Count >= 8)
                                                {
                                                    newExamRow.SubItems.Add(FileCtImg[4]);
                                                }
                                                else
                                                {
                                                    if (FileCtRep.Count >= 8)
                                                    { RelStat = " Rel: " + FileCtRep[4]; }
                                                    if (FileCtImg.Count >= 8)
                                                    { ImgStat = " Img: " + FileCtImg[4]; }
                                                    if (FileCtCpt.Count >= 8)
                                                    { CptStat = " Completo: " + FileCtCpt[4]; }
                                                    newExamRow.SubItems.Add(RelStat + ImgStat + CptStat);
                                                }

                                            }



                                        }
                                    }
                                    //Ficheiros
                                    if (FileCt[6] != "default")
                                    { newExamRow.SubItems.Add(FileCt[6]); ProcessedFiles.Add(FileCt[6]); }
                                    else { newExamRow.SubItems.Add(""); }
                                    if (FileCt[7] != "default" && FileCt[7] != "false")
                                    { newExamRow.SubItems.Add(FileCt[7]); ProcessedFiles.Add(FileCt[7]); }
                                    else
                                    { newExamRow.SubItems.Add(""); }
                                    if (FileCt[8] != "default")
                                    { newExamRow.SubItems.Add(FileCt[8]); ProcessedFiles.Add(FileCt[8]); }
                                    else { newExamRow.SubItems.Add(""); }




                                    
                                    if ((FileCt[6] != "default") && (FileCt[7] != "default"))
                                    {
                                        newExamRow.BackColor = Color.LightGreen;
                                    }
                                    else { newExamRow.BackColor = Color.LightYellow; }

                                    if (MyHelper.FileInUse(a.FullName.Replace(".txt", ".pdf")))
                                    { newExamRow.BackColor = Color.IndianRed; }

                                    if (dontjoin == false)
                                    {
                                        if (newlist.Contains(newExamRow) == false)
                                        {
                                            newlist.Add(newExamRow);
                                        }
                                    }
                                }
                            }
                        }
                               
                                    
                                
                            
                            
                        }
                        //comparar listas

                    List<string> newliststring = new List<string>();
                    foreach (ListViewItem c in newlist)
                    {
                        newliststring.Add(c.SubItems[0].Text + c.SubItems[1].Text + c.SubItems[2].Text + c.SubItems[3].Text + c.SubItems[4].Text + c.SubItems[5].Text + c.SubItems[6].Text + c.SubItems[7].Text + c.SubItems[8].Text);
                    }
                    var areEquivalent = (newliststring.Count() == oldlist.Count()) && !newliststring.Except(oldlist).Any();
                    
                    if (areEquivalent == false)
                    {
                        object[] parameters2 = new object[] { newlist };
                        BCKUpdateLst.ReportProgress(i, parameters2);
                        //reportar nova lista                        
                    }   

                    }
                


            }
            
            if (erros.Count > 0)
            {
                internallogger.AddEx(erros);
            }
             */

            #endregion
        }
        private void BCKUpdateLst_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 0)
            {
                try
                {
                    this.treeViewECG.BeginUpdate();
                    this.treeViewECG.Nodes.Clear();
                    object[] parameters = e.UserState as object[];
                    List<TreeNode> newTree = (List<TreeNode>)parameters[0];
                    this.treeViewECG.Nodes.AddRange(newTree.ToArray());
                    
                }
                catch { }
                finally { this.treeViewECG.EndUpdate(); }
                Gfirstload = false;
                gimediate = MyHelper.Formatmsg("Lista de ECG's atualizada.");
            }
            if (e.ProgressPercentage == 1)
            {
                try
                {
                    this.treeViewECO.BeginUpdate();
                    this.treeViewECO.Nodes.Clear();
                    object[] parameters = e.UserState as object[];
                    List<TreeNode> newTree = (List<TreeNode>)parameters[0];
                    this.treeViewECO.Nodes.AddRange(newTree.ToArray());

                }
                catch { }
                finally { this.treeViewECO.EndUpdate(); }
                
                Gfirstload = false;
                gimediate = MyHelper.Formatmsg("Lista de ECO's atualizada.");
            }
            if (e.ProgressPercentage == 2)
            {
                try
                {
                    this.treeViewHOLTER.BeginUpdate();
                    this.treeViewHOLTER.Nodes.Clear();
                    object[] parameters = e.UserState as object[];
                    List<TreeNode> newTree = (List<TreeNode>)parameters[0];
                    this.treeViewHOLTER.Nodes.AddRange(newTree.ToArray());

                }
                catch { }
                finally { this.treeViewHOLTER.EndUpdate(); }
                Gfirstload = false;
                gimediate = MyHelper.Formatmsg("Lista de HOLTER's atualizada.");
            }
            if (e.ProgressPercentage == 3)
            {
                try
                {
                    this.treeViewMAPA.BeginUpdate();
                    this.treeViewMAPA.Nodes.Clear();
                    object[] parameters = e.UserState as object[];
                    List<TreeNode> newTree = (List<TreeNode>)parameters[0];
                    this.treeViewMAPA.Nodes.AddRange(newTree.ToArray());

                }
                catch { }
                finally { this.treeViewMAPA.EndUpdate(); }
                Gfirstload = false;
                gimediate = MyHelper.Formatmsg("Lista de MAPA's atualizada.");
            }
            if (e.ProgressPercentage == 4)
            {
                try
                {
                    this.treeViewPE.BeginUpdate();
                    this.treeViewPE.Nodes.Clear();
                    object[] parameters = e.UserState as object[];
                    List<TreeNode> newTree = (List<TreeNode>)parameters[0];
                    this.treeViewPE.Nodes.AddRange(newTree.ToArray());

                }
                catch { }
                finally { this.treeViewPE.EndUpdate(); }
                Gfirstload = false;
                gimediate = MyHelper.Formatmsg("Lista de PE's atualizada.");
            }
            if (e.ProgressPercentage == 5)
            {
                try
                {
                    this.treeViewUC.BeginUpdate();
                    this.treeViewUC.Nodes.Clear();
                    object[] parameters = e.UserState as object[];
                    List<TreeNode> newTree = (List<TreeNode>)parameters[0];
                    this.treeViewUC.Nodes.AddRange(newTree.ToArray());

                }
                catch { }
                finally { this.treeViewUC.EndUpdate(); }
                Gfirstload = false;
                gimediate = MyHelper.Formatmsg("Lista de Não confirmados atualizada.");
            }
            #region OldUpdateList
            /*
            #region OldUpdateList
            if (e.ProgressPercentage == 0)
            {
                try
                {
                    object[] parameters = e.UserState as object[];
                    List<ListViewItem> newlist = (List<ListViewItem>)parameters[0];
                    lstECG.Items.Clear();
                    lstECG.Items.AddRange(newlist.ToArray());
                    //lstECG.Refresh();
                    tabPage1.Text = "ECG (" + newlist.Count.ToString() + ")";
                    gimediate = MyHelper.Formatmsg("Lista de ECG's atualizada.");
                }
                catch { }
            }
            if (e.ProgressPercentage == 1)
            {
                try
                {
                    object[] parameters = e.UserState as object[];
                    List<ListViewItem> newlist = (List<ListViewItem>)parameters[0];
                    lstECO.Items.Clear();
                    lstECO.Items.AddRange(newlist.ToArray());
                    // lstECO.Refresh();
                    tabPage4.Text = "ECO (" + newlist.Count.ToString() + ")";
                    gimediate = MyHelper.Formatmsg("Lista de ECO's atualizada.");
                }
                catch { }
            }
            if (e.ProgressPercentage == 2)
            {
                try
                {
                    object[] parameters = e.UserState as object[];
                    List<ListViewItem> newlist = (List<ListViewItem>)parameters[0];
                    lstHolter.Items.Clear();
                    lstHolter.Items.AddRange(newlist.ToArray());
                    //lstHolter.Refresh();
                    tabPage2.Text = "HOLTER (" + newlist.Count.ToString() + ")";
                    gimediate = MyHelper.Formatmsg("Lista de Holter's atualizada.");
                }
                catch { }
            }
            if (e.ProgressPercentage == 3)
            {
                try
                {
                    object[] parameters = e.UserState as object[];
                    List<ListViewItem> newlist = (List<ListViewItem>)parameters[0];
                    lstMAPA.Items.Clear();
                    lstMAPA.Items.AddRange(newlist.ToArray());
                    //lstMAPA.Refresh();
                    tabPage3.Text = "MAPA (" + newlist.Count.ToString() + ")";
                    gimediate = MyHelper.Formatmsg("Lista de MAPA's atualizada.");
                }
                catch { }
            }
            if (e.ProgressPercentage == 4)
            {
                try
                {
                    object[] parameters = e.UserState as object[];
                    List<ListViewItem> newlist = (List<ListViewItem>)parameters[0];
                    lstPE.Items.Clear();
                    lstPE.Items.AddRange(newlist.ToArray());
                    //lstPE.Refresh();
                    tabPage5.Text = "PE (" + newlist.Count.ToString() + ")";
                    gimediate = MyHelper.Formatmsg("Lista de PE's atualizada.");
                }
                catch { }
            }
            if (e.ProgressPercentage == 5)
            {
                try
                {
                    object[] parameters = e.UserState as object[];
                    List<ListViewItem> newlist = (List<ListViewItem>)parameters[0];
                    lstUnconfirmed.Items.Clear();
                    lstUnconfirmed.Items.AddRange(newlist.ToArray());
                    //lstPE.Refresh();
                    tabPage6.Text = "Não confirmados (" + newlist.Count.ToString() + ")";
                    gimediate = MyHelper.Formatmsg("Lista de não confirmados atualizada.");
                }
                catch { }
            } 
            #endregion*/

            #endregion
        }
        private void BCKEmail_DoWork(object sender, DoWorkEventArgs e)
        {
            List<string> erros = new List<string>();
            int error = 0;
            Helper internalhelp = new Helper();
            object[] parameters = e.Argument as object[];
            ListLogger internallogger = (ListLogger)parameters[0];
            string pathPDFOut = "";
            string PDFPrinter = "";
            string EmailPath = "";
            string EmailUser = "";
            string EmailPass = "";
            string EmailSender = "";
            string ExternalProc = "";
            string ExternalAppPath = "";
            string EmailReceiver = "";
            try { pathPDFOut = (string)parameters[1];}
            catch { pathPDFOut = ""; erros.Add (internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a pathPDFOut em BCKEmail."));}
            try { PDFPrinter = (string)parameters[2];}
            catch { PDFPrinter = ""; error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a PDFPrinter em BCKEmail.")); }
            try { EmailPath = (string)parameters[3];}
            catch { EmailPath = ""; error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a EmailPath em BCKEmail.")); }
            try { EmailUser = (string)parameters[4];}
            catch { EmailUser = ""; error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a EmailUser em BCKEmail.")); }
            try { EmailPass = (string)parameters[5];}
            catch { EmailPass = ""; error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a EmailPass em BCKEmail.")); }
            try { EmailSender = (string)parameters[8];}
            catch { EmailSender = ""; error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a EmailSender em BCKEmail.")); }
            try { ExternalProc = (string)parameters[6];}
            catch { ExternalProc = ""; error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a ExternalProc em BCKEmail.")); }
            try { ExternalAppPath = (string)parameters[7];}
            catch { ExternalAppPath = ""; error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a ExternalAppPath em BCKEmail.")); }
            try { EmailReceiver = (string)parameters[9]; }
            catch { EmailReceiver = ""; error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a EmailReceiver em BCKEmail.")); }
            if (error == 0)
            {
                
                if (pathPDFOut != "" && PDFPrinter != "" && EmailPath != "" && EmailUser != "" && EmailPass != "" && EmailSender != "" && ExternalProc != "" && ExternalAppPath != "" && EmailReceiver != "")
                {
                    
                   
                    if (Directory.Exists(pathPDFOut) && Directory.Exists(EmailPath) && internalhelp.printerisvalid(PDFPrinter) && File.Exists (ExternalAppPath) == true)
                    {
                        
                        try {
                            internalhelp.EmailCheck(EmailUser, EmailPass, EmailSender, EmailReceiver, EmailPath);
                            Thread.Sleep(5000);

                        }
                        catch { erros.Add(internalhelp.Formatmsg("Ocorreu um erro em EmailCheck em BCKEmail."));}
                        DirectoryInfo di = new DirectoryInfo(EmailPath);
                        List<FileInfo> PdfFiles = di.GetFiles("*.pdf", SearchOption.TopDirectoryOnly).ToList();
                        foreach (FileInfo Pdf in PdfFiles)
                        {
                            if (internalhelp.FileInUse(Pdf.FullName) == false)
                            {
                                try { Pdf.MoveTo(pathPDFOut + @"\" + Pdf.Name); }
                                catch { erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao mover " +  Pdf.Name+ " em BCKEmail.")); }
                            }
                        }
                        
                        if (internalhelp.CanIPrintHolters() == true)
                        {
                            
                            List<FileInfo> RPSFiles = di.GetFiles("*.rps", SearchOption.TopDirectoryOnly).ToList();
                            foreach (FileInfo RPS in RPSFiles)
                            {
                                if (internalhelp.FileInUse(RPS.FullName) == false)
                                {
                                    if (File.Exists(RPS.FullName.Replace(".RPS", ".txt")) == false)
                                    {
                                        File.WriteAllText(RPS.FullName.Replace(".RPS", ".txt"), "Start");
                                        Thread.Sleep(1000);
                                    }
                                    
                                    if (File.Exists(RPS.FullName.Replace(".RPS", ".txt")) == true)
                                    {
                                        List<string> RPSState = File.ReadAllLines(RPS.FullName.Replace(".RPS", ".txt")).ToList();
                                        if (RPSState.Count > 0)
                                        {
                                            if (RPSState[0] == "Finished." || RPSState[0] == "Ficheiro danificado.")
                                            { try { File.Delete(RPS.FullName.Replace(".RPS", ".txt")); 
                                            File.Delete(RPS.FullName);
                                            } catch { } }
                                            if (RPSState[0] == "Start")
                                            { internalhelp.PrintRPS (RPS.FullName ,ExternalAppPath,PDFPrinter,RPS.FullName.Replace(".RPS", ".txt"));}
                                        }
                                        else { try { File.Delete(RPS.FullName.Replace(".RPS", ".txt")); } catch { } }
                                    
                                    }
                                }
                            }
                            foreach (FileInfo RPS in RPSFiles)
                            {
                                if (File.Exists(RPS.FullName.Replace(".RPS", ".txt")) == true && (File.Exists(RPS.FullName) == true))
                                {
                                        List<string> RPSState = File.ReadAllLines(RPS.FullName.Replace(".RPS", ".txt")).ToList();
                                        if (RPSState.Count > 0)
                                        {
                                            if (RPSState[0] == "Finished." || RPSState[0] == "Ficheiro danificado.")
                                            { try { File.Delete(RPS.FullName.Replace(".RPS", ".txt")); 
                                            File.Delete(RPS.FullName);
                                            } catch { } }
                                            
                                        }
                                        else { try { File.Delete(RPS.FullName.Replace(".RPS", ".txt")); } catch { } }
                                }
                            }
                        }
                    }
                }
            }
            if (erros.Count > 0)
            {
                internallogger.AddEx(erros);
            }
        }
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            
        }
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            
        }
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            
        }
        private void contextMenuStrip4_Click(object sender, EventArgs e)
        {
            
        }
        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            
        }
        public void RunPDFOutPut(object[] parameters)
        {

            BCKPDFOutput.RunWorkerAsync(parameters);
        }
        public void RunFolders(object[] parameters)
        {

            BCKFolders.RunWorkerAsync(parameters);
        }
        public void RunEmail(object[] parameters)
        {
            if (BCKEmail.IsBusy == false)
            {
                BCKEmail.RunWorkerAsync(parameters);
            }
        }
        public void RunReport(object[] parameters)
        {

            BckReport.RunWorkerAsync(parameters);
        }
        private void lstHolter_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            
        }
        private void lstECG_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        private void lstHolter_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        private void lstMAPA_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        private void lstECO_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        private void lstPE_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        private void mudarEstadoDeValidaçãoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        private void mudarEstadoDeValidaçãoToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            
        }
        private void mudarEstaoDeValidaçãoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        private void mudaEstadoDeValidaçãoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        private void contextMenuStrip5_Click(object sender, EventArgs e)
        {
            
        }

        private void BckReport_DoWork(object sender, DoWorkEventArgs e)
        {
            List<string> erros = new List<string>();
            int error = 0;
            Helper internalhelp = new Helper();
            object[] parameters = e.Argument as object[];
            ListLogger internallogger = (ListLogger)parameters[0];
            string EmailUser = "";
            string EmailPass = "";
            string EmailDestination = "";
            string CompletePath = "";
          
            try { EmailUser = (string)parameters[1]; }
            catch { EmailUser = ""; error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a EmailUser em BckReport.")); }
            try { EmailPass = (string)parameters[2]; }
            catch { EmailPass = ""; error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a EmailPass em BckReport.")); }
            try { EmailDestination = (string)parameters[3]; }
            catch { EmailDestination = ""; error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a EmailDestination em BckReport.")); }
            try { CompletePath = (string)parameters[4]; }
            catch { CompletePath = ""; error++; erros.Add(internalhelp.Formatmsg("Ocorreu um erro ao atribuir valor a CompletePath em BckReport.")); }

            
            List<DirectoryInfo> subDirectories = new List<DirectoryInfo>();
            if (Directory.Exists(CompletePath))
            {
                List<string> di = Directory.GetDirectories(CompletePath).ToList();
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

                        List<string> Files = Directory.GetFiles(sd.FullName, "HOLTER.C.*.pdf").ToList();
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
                
                File.WriteAllLines(@"C:\PriJobAide\Logs\HolterRep.txt", Message.ToArray());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewECG;
            string path = gpathECG;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    string state = current.SelectedNode.Nodes[1].Text;
                    FrmContext action = new FrmContext(id, path, file);
                    action.ShowDialog();

                }
               
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewECG;
            string path = gpathECG;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    if (File.Exists(path + @"\" + file + ".txt"))
                    {
                        try {
                            List<string> oldcontent = File.ReadAllLines(path + @"\" + file + ".txt").ToList();
                            oldcontent[4] = "reset";
                            File.WriteAllLines(path + @"\" + file + ".txt", oldcontent.ToArray());
                        }
                        catch { }
                    }
                    

                }

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewECG;
            string path = gpathECG;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    if (File.Exists(path + @"\" + file + ".txt"))
                    {
                        try
                        {
                            List<string> oldcontent = File.ReadAllLines(path + @"\" + file + ".txt").ToList();
                            oldcontent[4] = "Apagar";
                            File.WriteAllLines(path + @"\" + file + ".txt", oldcontent.ToArray());
                        }
                        catch { }
                    }


                }

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewHOLTER;
            string path = gpathHOLTER;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    string state = current.SelectedNode.Nodes[1].Text;
                    FrmContext action = new FrmContext(id, path, file);
                    action.ShowDialog();

                }

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewHOLTER;
            string path = gpathHOLTER;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    if (File.Exists(path + @"\" + file + ".txt"))
                    {
                        try
                        {
                            List<string> oldcontent = File.ReadAllLines(path + @"\" + file + ".txt").ToList();
                            oldcontent[4] = "reset";
                            File.WriteAllLines(path + @"\" + file + ".txt", oldcontent.ToArray());
                        }
                        catch { }
                    }


                }

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewHOLTER;
            string path = gpathHOLTER;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    if (File.Exists(path + @"\" + file + ".txt"))
                    {
                        try
                        {
                            List<string> oldcontent = File.ReadAllLines(path + @"\" + file + ".txt").ToList();
                            oldcontent[4] = "Apagar";
                            File.WriteAllLines(path + @"\" + file + ".txt", oldcontent.ToArray());
                        }
                        catch { }
                    }


                }

            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewMAPA;
            string path = gpathMAPA;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    string state = current.SelectedNode.Nodes[1].Text;
                    FrmContext action = new FrmContext(id, path, file);
                    action.ShowDialog();

                }

            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewMAPA;
            string path = gpathMAPA;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    if (File.Exists(path + @"\" + file + ".txt"))
                    {
                        try
                        {
                            List<string> oldcontent = File.ReadAllLines(path + @"\" + file + ".txt").ToList();
                            oldcontent[4] = "reset";
                            File.WriteAllLines(path + @"\" + file + ".txt", oldcontent.ToArray());
                        }
                        catch { }
                    }


                }

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewMAPA;
            string path = gpathMAPA;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    if (File.Exists(path + @"\" + file + ".txt"))
                    {
                        try
                        {
                            List<string> oldcontent = File.ReadAllLines(path + @"\" + file + ".txt").ToList();
                            oldcontent[4] = "Apagar";
                            File.WriteAllLines(path + @"\" + file + ".txt", oldcontent.ToArray());
                        }
                        catch { }
                    }


                }

            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewECO;
            string path = gpathECO;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    string state = current.SelectedNode.Nodes[1].Text;
                    FrmContext action = new FrmContext(id, path, file);
                    action.ShowDialog();

                }

            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewECO;
            string path = gpathECO;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    if (File.Exists(path + @"\" + file + ".txt"))
                    {
                        try
                        {
                            List<string> oldcontent = File.ReadAllLines(path + @"\" + file + ".txt").ToList();
                            oldcontent[4] = "reset";
                            File.WriteAllLines(path + @"\" + file + ".txt", oldcontent.ToArray());
                        }
                        catch { }
                    }


                }

            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewECO;
            string path = gpathECO;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    if (File.Exists(path + @"\" + file + ".txt"))
                    {
                        try
                        {
                            List<string> oldcontent = File.ReadAllLines(path + @"\" + file + ".txt").ToList();
                            oldcontent[4] = "Apagar";
                            File.WriteAllLines(path + @"\" + file + ".txt", oldcontent.ToArray());
                        }
                        catch { }
                    }


                }

            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewPE;
            string path = gpathPE;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    string state = current.SelectedNode.Nodes[1].Text;
                    FrmContext action = new FrmContext(id, path, file);
                    action.ShowDialog();

                }

            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewPE;
            string path = gpathPE;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    if (File.Exists(path + @"\" + file + ".txt"))
                    {
                        try
                        {
                            List<string> oldcontent = File.ReadAllLines(path + @"\" + file + ".txt").ToList();
                            oldcontent[4] = "reset";
                            File.WriteAllLines(path + @"\" + file + ".txt", oldcontent.ToArray());
                        }
                        catch { }
                    }


                }

            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            TreeView current = this.treeViewPE;
            string path = gpathPE;
            if (current.SelectedNode != null)
            {
                if (current.SelectedNode.Level == 2 && current.SelectedNode.Nodes.Count >= 2)
                {
                    string id = current.SelectedNode.Nodes[0].Text;
                    string file = current.SelectedNode.Text;
                    if (File.Exists(path + @"\" + file + ".txt"))
                    {
                        try
                        {
                            List<string> oldcontent = File.ReadAllLines(path + @"\" + file + ".txt").ToList();
                            oldcontent[4] = "Apagar";
                            File.WriteAllLines(path + @"\" + file + ".txt", oldcontent.ToArray());
                        }
                        catch { }
                    }


                }

            }

        }

        
    }
}
