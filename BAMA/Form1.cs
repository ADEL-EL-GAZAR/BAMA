using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using BAMA.App_Code;

namespace BAMA
{
    public partial class FormBAMA : Form
    {
        public List<dictPrefixes> lst_dictPrefixes;
        public List<dictStems> lst_dictStems;
        public List<dictSuffixes> lst_dictSuffixes;
        public List<tableab> lst_tableab;
        public List<tableac> lst_tableac;
        public List<tablebc> lst_tablebc;
        List<CompatibilityTable> lst_Compatibility = new List<CompatibilityTable>();

        public FormBAMA()
        {
            InitializeComponent();
        }

        private void FormBAMA_Load(object sender, EventArgs e)
        {
            lst_dictPrefixes = Read_dictPrefixes();
            lst_dictStems = Read_dictStems();
            lst_dictSuffixes = Read_dictSuffixes();
            lst_tableab = Read_tableab();
            lst_tableac = Read_tableac();
            lst_tablebc = Read_tablebc();
        }

        private void btnAnalysis_Click(object sender, EventArgs e)
        {
            string word = txtWord.Text;
            if (word.Length > 0)
            {
                for (int pre = 0; pre < 5; pre++)
                {
                    string prefix = word.Substring(0, pre);
                    string stem_suffix = word.Substring(pre, word.Length - pre);
                    if (stem_suffix.Length > 1)
                    {
                        for (int suf = 0; suf < 7; suf++)
                        {
                            if (stem_suffix.Length - suf > 0)
                            {
                                string stem = stem_suffix.Substring(0, stem_suffix.Length - suf);
                                string suffix = stem_suffix.Substring(stem_suffix.Length - suf, suf);

                                if (
                                    lst_dictPrefixes.Where(i => i.A == prefix).FirstOrDefault() != null && 
                                    lst_dictStems.Where(i => i.A == stem).FirstOrDefault() != null && 
                                    lst_dictSuffixes.Where(i => i.A == suffix).FirstOrDefault() != null &&
                                    lst_tableab.Contains(new tableab(prefix, stem)) && 
                                    lst_tableac.Contains(new tableac(prefix, suffix)) && 
                                    lst_tablebc.Contains(new tablebc(stem, suffix)))
                                {
                                    dictPrefixes prefix_row = lst_dictPrefixes.Where(i => i.A == prefix).FirstOrDefault();
                                    dictStems stem_row = lst_dictStems.Where(i => i.A == stem).FirstOrDefault();
                                    dictSuffixes suffix_row = lst_dictSuffixes.Where(i => i.A == suffix).FirstOrDefault();

                                    MessageBox.Show(
                                        prefix_row.A + "\t" + stem_row.A + "\t" + suffix_row.A + "\t\n" +
                                        prefix_row.B + "\t" + stem_row.B + "\t" + suffix_row.B + "\t\n" +
                                        prefix_row.C + "\t" + stem_row.C + "\t" + suffix_row.C + "\t\n" +
                                        prefix_row.D + "\t" + stem_row.D + "\t" + suffix_row.D + "\t\n" +
                                        prefix_row.E + "\t" + stem_row.E + "\t" + suffix_row.E + "\t\n");
                                }
                            }
                        }
                    }
                }
            }

        }

        private List<dictPrefixes> Read_dictPrefixes()
        {
            List<dictPrefixes> lst = new List<dictPrefixes>();
            string appPath = Path.GetDirectoryName(Application.ExecutablePath);

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(appPath + "\\Resources\\dictPrefixes.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (int row = 1; row <= range.Rows.Count; row++)
            {
                lst.Add(new dictPrefixes(
                    (string)(range.Cells[row, 1] as Excel.Range).Value2,
                    (string)(range.Cells[row, 2] as Excel.Range).Value2,
                    (string)(range.Cells[row, 3] as Excel.Range).Value2,
                    (string)(range.Cells[row, 4] as Excel.Range).Value2,
                    (string)(range.Cells[row, 5] as Excel.Range).Value2));

            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            return lst;
        }

        private List<dictStems> Read_dictStems()
        {
            List<dictStems> lst = new List<dictStems>();
            string appPath = Path.GetDirectoryName(Application.ExecutablePath);

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(appPath + "\\Resources\\dictStems.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (int row = 1; row <= range.Rows.Count; row++)
            {
                lst.Add(new dictStems(
                    (string)(range.Cells[row, 1] as Excel.Range).Value2,
                    (string)(range.Cells[row, 2] as Excel.Range).Value2,
                    (string)(range.Cells[row, 3] as Excel.Range).Value2,
                    (string)(range.Cells[row, 4] as Excel.Range).Value2,
                    (string)(range.Cells[row, 5] as Excel.Range).Value2));

            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            return lst;
        }

        private List<dictSuffixes> Read_dictSuffixes()
        {
            List<dictSuffixes> lst = new List<dictSuffixes>();
            string appPath = Path.GetDirectoryName(Application.ExecutablePath);

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(appPath + "\\Resources\\dictSuffixes.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (int row = 1; row <= range.Rows.Count; row++)
            {
                lst.Add(new dictSuffixes(
                    (string)(range.Cells[row, 1] as Excel.Range).Value2,
                    (string)(range.Cells[row, 2] as Excel.Range).Value2,
                    (string)(range.Cells[row, 3] as Excel.Range).Value2,
                    (string)(range.Cells[row, 4] as Excel.Range).Value2,
                    (string)(range.Cells[row, 5] as Excel.Range).Value2));

            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            return lst;
        }

        private List<tableab> Read_tableab()
        {
            List<tableab> lst = new List<tableab>();
            string appPath = Path.GetDirectoryName(Application.ExecutablePath);

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(appPath + "\\Resources\\tableab.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (int row = 1; row <= range.Rows.Count; row++)
            {
                lst.Add(new tableab(
                    (string)(range.Cells[row, 1] as Excel.Range).Value2,
                    (string)(range.Cells[row, 2] as Excel.Range).Value2));

            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            return lst;
        }

        private List<tableac> Read_tableac()
        {
            List<tableac> lst = new List<tableac>();
            string appPath = Path.GetDirectoryName(Application.ExecutablePath);

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(appPath + "\\Resources\\tableac.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (int row = 1; row <= range.Rows.Count; row++)
            {
                lst.Add(new tableac(
                    (string)(range.Cells[row, 1] as Excel.Range).Value2,
                    (string)(range.Cells[row, 2] as Excel.Range).Value2));

            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            return lst;
        }

        private List<tablebc> Read_tablebc()
        {
            List<tablebc> lst = new List<tablebc>();
            string appPath = Path.GetDirectoryName(Application.ExecutablePath);

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(appPath + "\\Resources\\tablebc.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (int row = 1; row <= range.Rows.Count; row++)
            {
                lst.Add(new tablebc(
                    (string)(range.Cells[row, 1] as Excel.Range).Value2,
                    (string)(range.Cells[row, 2] as Excel.Range).Value2));

            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            return lst;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
