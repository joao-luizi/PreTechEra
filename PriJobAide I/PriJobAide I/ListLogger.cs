using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PriJobAide
{
    public class ListLogger
    {

        private List<string> Mylist = new List<string>();
        private object _sync = new object();
        public void AddEx(List<string> value)
        {
            lock (_sync)
            {
                Mylist.AddRange(value);
            }

        }
        public void RemoveEx(List<string> value)
        {
            lock (_sync)
            {
                Mylist = Mylist.Except(value).ToList();
            }

        }
        public int CountEx
        {
            get
            {
                lock (_sync)
                {
                    return Mylist.Count;
                }
            }
        }
        public List<string> GetList
        {
            get
            {
                lock (_sync)
                {
                    return Mylist;
                }
            }
        }

    }
    /*Usage example
     * private void MasterWriter_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] parameters = e.Argument as object[];
            ListLogger MWtowrite = (ListLogger)parameters[0];
            string InternalLogPath = (string)parameters[1];
            string file = InternalLogPath + DateTime.Now.ToShortDateString() + ".txt";
            if (File.Exists(file))
            {
                bool fileisinuse = helper.FileInUse(file);
                int error = 0;
                if (fileisinuse == true)
                {
                    
                    DateTime quit = DateTime.Now;
                    while (fileisinuse == true)
                    {                      
                        Thread.Sleep(500);
                        fileisinuse = helper.FileInUse(file);
                        if (quit.AddSeconds(40) < DateTime.Now)
                        {
                            error++;
                            break; 
                        }
                    }
                }
                if (error == 0)
                {
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
                using (StreamWriter sw = File.CreateText(file))
                {
                    MWtowrite.GetList.ForEach(sw.WriteLine);
                    sw.Close();
                }
                MWtowrite.RemoveEx(MWtowrite.GetList);

            }                                                          
        }
     * */

}
