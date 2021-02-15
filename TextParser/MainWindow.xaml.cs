using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Microsoft.Win32;

namespace TextParser
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    /// 

    public partial class MainWindow : Window
    {
        public string text_to_parse;
        public string path_to_text;
        public Dictionary<string, int> parsed_text;
        // cancellation token for worker task
        public static CancellationTokenSource cts = new CancellationTokenSource();
        
        public MainWindow()
        {
            InitializeComponent();
        }

        // ========================= EVENTS ==========================================
        private void btn_from_path_Click(object sender, RoutedEventArgs e)
        {
            //opening file browser to select target file
            //works with txt and rtf files
            OpenFileDialog o_dlg = new OpenFileDialog();
            o_dlg.DefaultExt = ".txt";

            Nullable<bool> result = o_dlg.ShowDialog();

            if (result == true)
            {
                string filename = o_dlg.FileName;
                txtbox_from_path.Text = filename;
                text_to_parse = System.IO.File.ReadAllText(@filename); //works!!
                if (txtbox_to_path.Text != "")
                {
                    btn_start.IsEnabled = true;
                }
            }
        }

        private void btn_start_Click(object sender, RoutedEventArgs e)
        {
            string mode = ""; 
            if (export_txt.IsChecked == true)
            {
                mode = "txt";
            }
            else if(export_excel.IsChecked == true)
            {
                mode = "excel";
            }
            //starting worker task for "heavy lifting" methods
            Task.Factory.StartNew(() => parse_text(text_to_parse, mode));
        }

        private void btn_cancel_Click(object sender, RoutedEventArgs e)
        {
            //sigalling the cancellation token to cancel
            cts.Cancel();
        }

        private void btn_to_path_Click(object sender, RoutedEventArgs e)
        {
            //opening file browser for saving location
            SaveFileDialog s_dlg = new SaveFileDialog();
            s_dlg.FileName = "untitled";

            if (export_txt.IsChecked == true)
            {
                s_dlg.Filter = "Text (*.txt)|*.txt";
            }
            else if (export_excel.IsChecked == true)
            {
                s_dlg.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            }

            Nullable<bool> result = s_dlg.ShowDialog();

            if (result == true)
            {
                string filename = s_dlg.FileName;
                txtbox_to_path.Text = filename;
                path_to_text = filename;
                if (txtbox_from_path.Text != "")
                {
                    btn_start.IsEnabled = true;
                }
            }
        }

        private void export_txt_Click(object sender, RoutedEventArgs e)
        {
            change_file_extension();
        }
        private void export_excel_Click(object sender, RoutedEventArgs e)
        {
            change_file_extension();
        }
        // ===================== METHODS ==================================================

        private void change_file_extension()
        {
            if (txtbox_to_path.Text != "")
            {
                string to_path = txtbox_to_path.Text;
                string[] path = to_path.Split('.');
                if (export_txt.IsChecked == true)
                {
                    path[1] = ".txt";
                }
                else if(export_excel.IsChecked == true)
                {
                    path[1] = ".xls";
                }
                txtbox_to_path.Text = path[0] + path[1];
                path_to_text = path[0] + path[1];
            }
        }
        //method for parsing 
        private void parse_text(string text_to_parse, string mode)
        {
            //splitting the string into words to parse them into a key/value dictionary
            CancellationToken token = cts.Token;
            bool canceled = false;
            Dictionary<string, int> dictionary = new Dictionary<string, int>();

            //setting the progress bar early to signal something is happening
            //using dispatcher to be able to modify control from another thread
            Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                progressbar.Value = 10;
            }));
            string[] words = text_to_parse.Split(new[] {" ", "\r\n", "\r", "\n" }, StringSplitOptions.None);
            int entries = words.Count();

            Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                progressbar.Value = 20;
            }));

            int one_percent = entries / 70;
            int percentage = 30;
            int a = 0;

            foreach (string word in words)
            {
                a++;
                if (a == one_percent)
                {
                    a = 0;
                    percentage++;
                    Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
                    {
                        progressbar.Value = percentage;
                    }));
                }

                if (dictionary.ContainsKey(word))
                {
                    dictionary[word] += 1;
                }
                else
                {
                    dictionary.Add(word, 1);
                }
                if (token.IsCancellationRequested)
                {
                    cts.Dispose();
                    cts = new CancellationTokenSource();
                    canceled = true;
                    break;
                }
            }
            Dictionary<string, int> sorted_dictionary = new Dictionary<string, int>();

            //sorting all entries descending 
            foreach (var item in dictionary.OrderByDescending(key => key.Value))
            {
                sorted_dictionary.Add(item.Key, item.Value);
            }

            if (canceled)
            {
                cts.Dispose();
                cts = new CancellationTokenSource();
                MessageBox.Show("Parsing to new file canceled", "Aborted");

            }
            else
            {
                if (mode == "txt")
                {
                    create_txt_file(path_to_text, sorted_dictionary);
                }
                else if (mode == "excel")
                {
                    create_excel_file(path_to_text, sorted_dictionary);
                }
                MessageBox.Show("Parsing to new file complete", "Finished");
            }
            
            Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                progressbar.Value = 0;
            }));
        }

        //method for creating the txt file from parsed text via file-stream
        private void create_txt_file(string filepath, Dictionary<string, int> parsed_text)
        {
            try
            {
                if (File.Exists(filepath))
                {
                    File.Delete(filepath);
                }

                using (FileStream filestream = File.Create(filepath))
                {
                    foreach (KeyValuePair<string, int> line in parsed_text)
                    {
                        //going line by line entering the dictionaries content into the .txt
                        Byte[] text = new UTF8Encoding(true).GetBytes(line.Key + " - " + line.Value + "\r\n");
                        filestream.Write(text, 0, text.Length);
                    }
                }
                using (StreamReader streamreader = File.OpenText(filepath))
                {
                    string s = "";
                    while ((s = streamreader.ReadLine()) != null)
                    {
                        Console.WriteLine(s);
                    }
                }
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex.ToString());
            }
        }

        public void create_excel_file(string filepath, Dictionary<string, int> parsed_text)
        {
            if (File.Exists(filepath))
            {
                File.Delete(filepath);
            }

            Microsoft.Office.Interop.Excel.Application xlApp = new
            Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Wort";
            xlWorkSheet.Cells[1, 2] = "Anzahl";

            int a = 2;
            foreach (KeyValuePair<string, int> line in parsed_text)
            {
                    xlWorkSheet.Cells[a, 1] = line.Key;
                    xlWorkSheet.Cells[a, 2] = line.Value;
                    a++;
            }

            //save as xls file!
            xlWorkBook.SaveAs(filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }


    }


}
