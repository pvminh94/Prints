using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
namespace Prints
{
    class Program
    {
        
        Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        Document document =null;
        FolderBrowserDialog fbd = new FolderBrowserDialog();
        var dlg = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();
        object missing = System.Reflection.Missing.Value;
        public void DisposeExcelInstance()
        {
           
        }
        void Process()
        {
          
            try
            {
                
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    foreach (string path in Directory.GetFiles(fbd.SelectedPath))
                    {
                       
                        object file = (object)path;
                            document = app.Documents.Open(ref file, ref missing, ref missing,
                                                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                        document.PrintOut();
                    }
                }
            }
            catch(Exception e)
            {
                e.ToString();
            }
            finally
            {
                //Marshal.ReleaseComObject(document);
               // Marshal.ReleaseComObject(app);
               document.Close(ref missing, ref missing, ref missing);
               app.Quit(ref missing, ref missing, ref missing);
            }
        }
        [STAThread]
        static void Main(string[] args)
        {
            Program p = new Program();
            p.DisposeExcelInstance();
            p.Process();
            MessageBox.Show("print complete!");
            Environment.Exit(0);
        }
    }
}
