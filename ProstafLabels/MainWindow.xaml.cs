using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ProstafLabels
{
    public partial class MainWindow : Window
    {
        public string excelPath = string.Empty;
        public string wordPath = string.Empty;

        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;

        private static Word.Application OriginalApp = null;
        private static Word.Document OriginalDoc = null;

        List<string> array = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void importWorksheet_Click(object sender, RoutedEventArgs e)
        {
            Thread importingsheets = new Thread(new ThreadStart(_importsheets));
            importingsheets.Start();
        }

        public void _importsheets()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                logList.Items.Add("Importing rows...");
                if (sheetList.SelectedItem != null)
                    for (int i = 1; i <= MyBook.Sheets.Count; i++)
                        if (MyBook.Sheets[i].Name == sheetList.SelectedItem.ToString())
                            for (int j = 1; j <= MyBook.Sheets[i].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; j++)
                                databaseList.Items.Add(
                                    MyBook.Sheets[i].Cells[j, 1].value + "," +
                                    MyBook.Sheets[i].Cells[j, 2].value + "," +
                                    MyBook.Sheets[i].Cells[j, 3].value + "," +
                                    MyBook.Sheets[i].Cells[j, 4].value);
            }));
        }

        private void printButton_Click(object sender, RoutedEventArgs e)
        {
            int k = databaseList.Items.Count / 24;
            OriginalApp = new Word.Application()
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };

            for (int j = 0; j < 2; j++)
            {
                OriginalDoc = OriginalApp.Documents.Open(wordPath);
                OriginalDoc.Activate();

                for (int i = 0; i < 24; i++)
                {
                    array = databaseList.Items[i].ToString().Split(',').ToList();
                    databaseList.Items.Remove(databaseList.Items[i]);

                    FindAndReplace(OriginalApp, "<name" + i + ">", array[0].ToUpper());
                    FindAndReplace(OriginalApp, "<address" + i + ">", array[1].ToUpper());
                    FindAndReplace(OriginalApp, "<postcode" + i + ">", array[2].ToUpper());
                    FindAndReplace(OriginalApp, "<city" + i + ">", array[3].ToUpper());
                }

                array.Clear();
                OriginalDoc.PrintOut(Background:false);
                OriginalDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            }
        }

        void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, string findText, string replaceWithText)
        {
            if (replaceWithText.Length > 255)
            {
                FindAndReplace(doc, findText, findText + replaceWithText.Substring(255));
                replaceWithText = replaceWithText.Substring(0, 255);
            }

            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            doc.Selection.Find.Execute(findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

        private void templateButton_Click(object sender, RoutedEventArgs e)
        {
            logList.Items.Add("Importing word file...");
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            dlg.DefaultExt = ".doc";

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                wordPath = dlg.FileName;
                logList.Items.Add("Finished");
            }
            else 
                logList.Items.Add("Failed");
        }

        private void fileButton_Click(object sender, RoutedEventArgs e)
        {
            Thread importingData = new Thread(new ThreadStart(_importingData));
            importingData.Start();
        } 

        public void _importingData()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                logList.Items.Add("Importing excel file...");
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

                dlg.DefaultExt = ".xlsx";
                dlg.Filter = "ExcelNames|*.xlsx";

                Nullable<bool> result = dlg.ShowDialog();

                if (result == true)
                {
                    excelPath = dlg.FileName;

                    MyApp = new Excel.Application
                    {
                        Visible = false
                    };

                    MyBook = MyApp.Workbooks.Open(excelPath);

                    for (int i = 1; i <= MyBook.Sheets.Count; i++)
                        sheetList.Items.Add(MyBook.Sheets[i].Name);

                    logList.Items.Add("Finished");
                }
                else
                    logList.Items.Add("Failed");
            }));
        }
    }
}
