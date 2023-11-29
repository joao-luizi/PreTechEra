using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;

namespace PriJobAide
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Main());
            
            bool debug = true;
            int state = 2;
            string PathIni = @"C:\PriJobAide\PriJobAide.ini";
            // To prevent the program to be started twice
            ///Create new mutex
            if (File.Exists(PathIni))
            {
                List<string> inivalues = new List<string>();
                inivalues = File.ReadAllLines(PathIni).ToList();
                if (inivalues.Count > 0)
                {
                    foreach (string s in inivalues)
                    {

                        if (s.Contains("debug="))
                        {
                            try
                            {
                                string[] split = s.Split('=');
                                debug = Convert.ToBoolean (split[1]);
                            }
                            catch { debug = true; }
                        }
                        if (s.Contains("state="))
                        {
                            try
                            {
                                string[] split = s.Split('=');
                                state = Convert.ToInt32(split[1]);
                            }
                            catch { state = 2; }
                        }
                    }
                }
                
                var exists = System.Diagnostics.Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly().Location)).Count() > 1;
                ///if creation of mutex is successful
                if (exists==false)
                {
                    
                    Application.Run(new Main(state,debug));
                    
                    
                }
                else
                {
                    
                    state = 2;
                    Application.Run(new Main(state,debug));
                    

                }
            }
            else
            {
                Application.Run(new FrmSetup());
            }
        }
    }
}
